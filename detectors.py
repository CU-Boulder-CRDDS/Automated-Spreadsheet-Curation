import pandas as pd # Reading excel filesh
import numpy as np # Numerical operations
import pprint # Pretty printing results for long output
from abc import ABC, abstractmethod # Creating abstract classes
import re # Regular expressions to test cell contents
import json # Saving results to json
from datetime import datetime as dt # Add timestamps to the results
import os # Operating system operations

class Test_Suite():
    __slots__ = ["wb_path", "wb", "to_run", "results"]

    # results is a dict of sheet names, each containing a dict of test names and their results
    # to_run is a list of test classes to run on each sheet
    # wb_path is the path to the workbook
    # wb is a dictionary of pandas dataframes, each representing a sheet in the workbook

    def __init__(self, wb_path, to_run = None):
        self.wb_path = wb_path

        # Initialize the results dict,
        # levels:  [sheet_name] -> [test_name] -> [test_object]
        self.results = dict()
        # Initialize the results dict for the file
        self.results["file"] = dict()

        # First check if the file is UTF-8 encoded
        encoding_test = File_Encoding(ws = None, wb_path = wb_path)
        encoding_test.validate()
        self.results["file"]["file_encoding"] = encoding_test


        # Extract file extension from wb_path
        file_extension = os.path.splitext(wb_path)[1]


        # If the file extension is .xlsx, read the file as an excel file
        if file_extension == ".xlsx":
            self.wb = pd.read_excel(
                io = wb_path, # The file path
                dtype = "str", # Parse all cells as strings, leaves the parsing to the tests
                sheet_name = None, # read all sheets returning a dict
                header = None, # Do not load the header at all
            )
        elif file_extension == ".csv":
            self.wb = pd.read_csv(
                wb_path, # The file path
                sep = None, # Automatically detect the separator
                dtype = "str", # Parse all cells as strings, leaves the parsing to the tests
                header = None, # Do not load the header at all
                na_values = "", # Treat only empty cells as NA
                keep_default_na = False, # Do not keep default NA values
                encoding = "utf-8", # Read the file as UTF-8 encoded
                encoding_errors = "replace", # Replace invalid characters with ?
            )
            # Make  the single sheet into a singleton dict
            self.wb = {os.path.splitext(os.path.basename(wb_path))[0]: self.wb}
        else:
            raise ValueError("Invalid file extension. Use '.xlsx' or '.csv'.")


        # Set the first row to be headers no matter what
        def set_headers(df):
            # If empty
            if df.empty:
                # Just return
                return df
            else:
                cols = df.iloc[0]
                cols = cols.mask(cols.isna(), "Unnamed")
                df.columns = cols
                df = df[1:]
                df = df.reset_index(drop=True)

                return df
        
        # Use set_headers on each sheet
        # Initiate new container
        
        # Iterate over dict
        if isinstance(self.wb, dict):
            # Get each dataframe in the dict
            self.wb = {name : set_headers(df) for name, df in self.wb.items()}
        else:
            self.wb = set_headers(self.wb)



        # If no tests are provided, run all tests
        if to_run == None:
            self.to_run = [Upper_Left_Corner, Multi_Table, Mixed_Datatypes, Header_Duplicates, Header_Length, Header_First_Char, Header_Space, Header_Word_Separation, Header_Special_Characters, Header_ID, Header_Date, Special_Char, Untrimmed_White_Space, Newlines_Tabs, Missing_Value_Text, Question_Mark_Only, White_Space_Only, Number_Space, Dates, Scientific_Notation]
        else:
            self.to_run = to_run






    def run(self):
        # Run tests on the file as a whole


        filename_tests = [Filename_Length, Filename_Whitespace, Filename_Final, Filename_Word_Separation, Filename_Special_Characters]

        for test in filename_tests:
            one_test = test(self.wb, self.wb_path)
            one_test.validate()
            self.results["file"][one_test.name] = one_test



        # For each sheet in the workbook
        for sheet in self.wb.keys():
            print(f"Running tests on {sheet}")

            # Get the sheet as a dataframe
            ws = self.wb[sheet]

            # First check if the sheet is empty
            # Initialize the empty test
            empty_test = Worksheet_Empty(ws, self.wb_path)

            # Validate the empty test
            empty_test.validate()

            # If the test failed
            if not empty_test.status:

                # the sheet is empty, skip it
                print(f"Sheet {sheet} is empty. Skipping.")
                continue


            # Initialize the results dict for this sheet
            self.results[sheet] = dict()

            # First check the worksheet name
            sheet_name_test = Worksheet_Name(ws, sheet)

            # Validate the sheet name test
            sheet_name_test.validate()

            # Add the results to the results
            self.results[sheet]["sheet_name"] = sheet_name_test

            # Initialize the queue of tests to run
            queue = self.to_run[:]

            # Initialize a dict to keep track of completed tests
            completed_tests = dict()

            while len(queue) > 0:
                    
                # Get the first test
                test_type = queue.pop(0)

                # Initialize the test
                test = test_type(ws)
                
                # Check if the test has dependencies
                if issubclass(test_type, Has_Dependency):

                    # Check if any dependencies are not completed in the 
                    # completed_tests dict
                    if any( 
                        map(
                            lambda depend: depend(ws).name not in completed_tests.keys(),
                            test.dependencies
                        ) 
                    ):
                        # If not, skip the test for now
                        # and add it back to the queue at the end
                        queue.append(test)

                        # Move to the next test
                        continue

                    # If all dependencies are completed, check their validity 
                    # to ensure they contain the needed information
                    else:
                        # Get a dict of the completed tests that are dependencies
                        fulfilled = {dep_name : dep_test for dep_name, dep_test in completed_tests.items() if type(dep_test) in test.dependencies}

                        # Provide these tests as arguments to the new test's
                        # validate method
                        test.validate(**fulfilled)
                else:
                    test.validate()
                
                # Add the completed test object to the results
                self.results[sheet][test.name] = test

                # Add the test to the dict of completed tests
                completed_tests[test.name] = test
                    



            
    # Print a report of the results
    def report(self):

        # For each sheet and its associated tests
        for sheet, tests in self.trimmed_results().items():
            # Add extra extra space between sheets
            print("\n\n" + 80 * "=")
            # Print the sheet name
            print(f"Sheet: {sheet}")
            # Add space
            print("\n")

            # For each tests applied to this sheet
            # trimmed_results is a dict of test names and their results
            for test_name, trimmed_result in tests.items():
                
                # Print the report for that test
                print(f"Test: {test_name}")
                print(f"Message: {pprint.pformat(trimmed_result['message'])}")
                print(f"Issues: {pprint.pformat(trimmed_result['issues'])}")

                # Add some space
                print("\n")

            

    # Trim the results to only include dict entries for failed tests
    def trimmed_results(self, stringify = False):
        
        # initialize the trimmed results dict
        trimmed_results = dict()

        # For each sheet and its associated tests
        for sheet, tests in self.results.items():


            # If any test has failed
            if not all([test.status for test in tests.values()]):

                # Add the sheet to the trimmed results
                trimmed_results[sheet] = dict()

                # For each test applied to this sheet
                for test_name, test in tests.items():

                    # If the test has failed
                    if not test.status:

                        # If strings needed instead of tuples
                        # This is because the keys are tuples of coordinates
                        # and cannot be saved as json
                        if stringify:
                            issues = {
                                str(key): val for key, val in test.issues.items()}
                        else:
                            issues = test.issues

                        # Add the test to the trimmed results
                        # with its status, message, and issues
                        trimmed_results[sheet][test_name] = {
                            "message": test.message,
                            "issues": issues
                        }
        
        # Return the trimmed results
        # This is a dict of sheet names, each containing a dict of test names and their results
        return trimmed_results

    # Save the results to a file
    def save(self, format="json", filename=None):
        import os
        from datetime import datetime as dt
        import json
        print("saving results to file")

        # By default, use "results/" as the folder
        results_folder = "results"
        if not os.path.exists(results_folder):
            os.makedirs(results_folder)

        # Get the base name of the file from wb_path (without folders)
        base_name = os.path.splitext(os.path.basename(self.wb_path))[0]
        # Add datetime and format to the output filename
        timestamp = dt.now().strftime("%Y-%m-%d_%H-%M-%S")
        default_filename = f"{base_name}_{timestamp}.{format}"

        # Full path to output file
        if filename is None:
            filename = os.path.join(results_folder, default_filename)

        if format == "json":
            with open(filename, "w") as f:
                json.dump(self.trimmed_results(stringify=True), f, default=str)
        elif format == "csv":

            rows = []
            # For each sheet and its associated tests
            for sheet, tests in self.trimmed_results().items():
                for test_name, test in tests.items():
                    issues = test["issues"]
                    if not issues:
                        # If no issues, optionally still record a "pass"/empty row
                        continue
                    row_count = len(issues)
                    # Prepare repeated/columnar values
                    paths     = [self.wb_path] * row_count
                    sheets    = [sheet] * row_count
                    test_names= [test_name] * row_count
                    messages  = [test["message"]] * row_count
                    locations, examples = zip(*issues.items())
                    for i in range(row_count):
                        rows.append({
                            "path": paths[i],
                            "sheet": sheets[i],
                            "test_name": test_names[i],
                            "message": messages[i],
                            "location": locations[i],
                            "example": examples[i]
                        })
            import pandas as pd  # ensure pd is imported
            df = pd.DataFrame(rows, columns=["path", "sheet", "test_name", "message", "location", "example"])
            df.to_csv(filename, mode="w", index=False, header=True)
        else:
            # Raise a ValueError
            raise ValueError("Invalid format. Use 'json' or 'csv'.")


# Abstract class for the structure of a test
class Test(ABC):


    def __init__(self, ws, name):
        # name of the test 
        self.name = name

        # Whether the test has been run
        self.is_run = False

        # True if the test passed, False if it failed
        self.status = None

        # The worksheet dataframe to run the test on
        self.ws = ws

        # The issues found during the test
        # This is a dict of cell coordinates and their associated issues
        self.issues = dict()

        # The message to display with the test results
        # This is a string that describes the outcome of the test in plain english
        self.message = "Not yet run"

    # The validate method is the main method that runs the test
    # Every Test class must implement this method
    # Most tests will use the status code given below at the end
    @abstractmethod
    def validate(self, fail_message, pass_message):

        # The test has been run
        self.is_run = True

        # If there were any issues
        if len(self.issues) > 0:

            # The test failed
            self.status = False

            # Save the fail message
            self.message = fail_message 
        else:

            # The test passed
            self.status = True

            # Save the pass message
            self.message = pass_message       

    # The report method prints the results of the test
    # This method is not currently used anywhere
    def report(self):
        print(f"name: {self.name}\nstatus: {self.status}\nmessage: {self.message}\nissues: {self.issues}")

    def handle_multi_table(self):
        Test.validate(
            self,
            None,
            "Cannot assess test if multiple tables are present.  Pass by default")



# Abstract class for tests that depend on other tests
class Has_Dependency(ABC):

    # The dependencies are provided as arguments 
    # Each dependency is a test class, not an instance
    def __init__(self, *dependencies):
        self.dependencies = dependencies

    # At the time of validation, instances of the dependencies are passed
    # We check that they have been run. 
    def check_input(self, *inputs):

        # Save the inputs
        self.inputs = inputs

        # Check that all dependencies have been run
        for test in self.inputs:

            # Raise an assertion error if the test has not been run
            # This should only ever happen if there is a bug in the code.
            # Nothing the user does can produce this issue.
            assert test.is_run, f"{test.name} dependency not yet run"

# Abstract class for tests that check each cell in the worksheet
class By_Cell(ABC):

    # Generator over pandas dataframes
    # tuple of the cell coordinates and the cell contents
    # This is used to iterate over the cells in the worksheet
    def pandas_iter(self):
        # For each row in the worksheet
        for row_idx, row in self.ws.iterrows():

            # For each column in the row
            for col_name, cell in row.items():

                # Yield the cell coordinates and the cell contents
                # (row_idx, (col_idx, col_name)), contents)
                yield (row_idx, (self.ws.columns.get_loc(col_name), col_name)), cell


    # Cell by cell validation.  
    # Assumes that the not_valid method is implemented
    def validate(self, fail_message, pass_message, df = None):

        # If provided, df may be a subset of the whole worksheet
        # If not provided, use the whole worksheet
        if df is None:
            df = self.ws

        
        for cell in self.pandas_iter():

            # If the cell is not valid, as given by subclass
            if self.not_valid(cell[1]):

                # Save the cell coordinates and the cell contents
                self.issues[cell[0]] = cell[1]
        

        # Set the status messages
        Test.validate(self, fail_message, pass_message)


    # The not_valid method is an abstract method that must be implemented
    # by each test class that inherits from By_Cell
    # This method should return True if the cell is not valid, and False if it is valid
    # The validate method will call this method for each cell in the worksheet
    @abstractmethod
    def not_valid(self, cell):
        pass

class Worksheet_Empty(Test):
    def __init__(self, ws, wb_path):
        # Save the workbook path
        self.wb_path = wb_path

        # Initialize the test named empty
        Test.__init__(self, ws, "empty")

    def validate(self):

        # Check if the worksheet is empty
        if self.ws.empty:
            # The test failed
            self.status = False

            # Set the message
            self.message = "Worksheet is empty"

            self.issues = {"filename": self.wb_path}

        else:
            # The test passed
            self.status = True

            # Set the message
            self.message = "Worksheet is not empty"

class Worksheet_Name(Test):
    def __init__(self, ws, sheet_name):
        # Save the sheet name
        self.sheet_name = sheet_name

        # Initialize the test named sheet_name
        Test.__init__(self, ws, "sheet_name")

    def validate(self):


        # detect special characters in filename
        spec_char = re.search(r"[!@#$%^&*()+=\[\]{};:'\"|\\,<>\?/]", self.sheet_name)
        if spec_char:
            self.status = False

            self.issues["contains special characters"] = spec_char.group(0)
        
        # detect whitespace in filename
        if re.search(r"\s", self.sheet_name):
            self.status = False

            self.issues["contains whitespace"] = self.sheet_name

        
        if  self.status is None:
            # The test passed
            self.status = True

            # Set the message
            self.message = "Worksheet name is acceptable"
        else:
            self.message = "Worksheet name is not acceptable"


class Filename(Test):
    def __init__(self, ws, wb_path, test_name):
        # Save the workbook path
        self.wb_path = wb_path
        # Save the filename
        self.filename = os.path.basename(self.wb_path)

        # Initialize the test named test_name
        Test.__init__(self, ws, test_name)
    
    def validate(self):
        pass

        
class Filename_Length(Filename):
    def __init__(self, ws, wb_path):
        # Initialize the test named filename_length
        Filename.__init__(self, ws, wb_path, "filename_length")

    def validate(self):

        # If length is greater than 200
        if len(self.filename) > 200:
            # The test failed
            self.status = False
            # Add an issue "location = filename": "contents of filename"
            self.issues["filename"] = self.filename
            # Set message
            self.message = "Filename is > 200 characters"
        else:
            # The test passed
            self.status = True
            # Set message
            self.message = "Filename is <= 200 characters"


class Filename_Whitespace(Filename):
    def __init__(self, ws, wb_path):
        # Initialize the test named filename_whitespace
        Filename.__init__(self, ws, wb_path, "filename_whitespace")

    def validate(self):

        # detect whitespace in filename
        # search returns None if no match is found, else Match object
        if re.search(r"\s", self.filename) is not None:

            # The test failed
            self.status = False

            # Add an issue
            self.issues["filename"] = self.filename

class Filename_Final(Filename):
    def __init__(self, ws, wb_path):
        # Initialize the test named filename_final
        Filename.__init__(self, ws, wb_path, "filename_final")

    def validate(self):

        # If filename contains 'final'
        if "final" in self.filename.lower():
            self.status = False

            self.issues["filename"] = self.filename

            self.message = "Filename contains 'final'"
        else:
            self.status = True
            self.message = "Filename does not contain 'final'"
    
class Filename_Word_Separation(Filename):
    def __init__(self, ws, wb_path):
        # Initialize the test named filename_word_separation
        Filename.__init__(self, ws, wb_path, "filename_word_separation")

    def validate(self):

        # Detect combo of camel case, underscore, and dash in filename
        # Expressions work even if there are whitespace in the filename
        # camel case: captial letters in the middle of words
        camel_case = re.search(r'[a-z][A-Z]', self.filename) is not None
        underscore = "_" in self.filename
        dash = "-" in self.filename

        # If more than one are true
        if sum([camel_case, underscore, dash]) > 1:
            self.status = False
            self.issues["filename"] = self.filename
            self.message = "Filename mixes camel case, underscores, and dashes"
        else:
            self.status = True
            self.message = "Filename does not mix camel case, underscores, and dashes"
        
class Filename_Special_Characters(Filename):
    def __init__(self, ws, wb_path):
        # Initialize the test named filename_special_characters
        Filename.__init__(self, ws, wb_path, "filename_special_characters")

    def validate(self):

        # detect special characters in filename
        spec_char = re.search(r"[!@#$%^&*()+=\[\]{};:'\"|\\,<>\?/]", self.filename)
        
        # If there was a special character
        if spec_char:
            # The test failed
            self.status = False

            # Add an issue
            self.issues["filename"] = spec_char.group(0)

            self.message = "Filename contains special characters"
        else:
            self.status = True
            self.message = "Filename does not contain special characters"
        

class File_Encoding(Test):
    def __init__(self, ws, wb_path):
        # Save the workbook path
        self.wb_path = wb_path
        # Initialize the test named file_encoding
        Test.__init__(self, ws, "file_encoding")
    
    def validate(self):
        # Extract file extension
        file_extension = os.path.splitext(self.wb_path)[1].lower()
        
        # Only check encoding for text-based files (CSV, not Excel)
        if file_extension == ".xlsx" or file_extension == ".xls":
            # Excel files are binary ZIP archives, encoding check doesn't apply
            self.status = True
            self.message = "Encoding check skipped for Excel files (binary format)"
            self.is_run = True
            return
        
        # For CSV and other text files, check UTF-8 encoding
        try:
            with open(self.wb_path, 'rb') as f:
                # Read the file content
                content = f.read()
                # Try to decode as UTF-8
                content.decode('utf-8')
            
            # If decoding succeeds, the test passed
            self.status = True
            self.message = "File encoding is UTF-8"
        except UnicodeDecodeError:
            # If decoding fails, the test failed
            self.status = False
            self.issues["file"] = "File encoding is not UTF-8"
            self.message = "File encoding is not UTF-8"
        except Exception as e:
            # Handle other exceptions (e.g., file not found)
            self.status = False
            self.issues["file"] = f"Error reading file: {str(e)}"
            self.message = f"Error checking file encoding: {str(e)}"
        
        # Mark the test as run
        self.is_run = True


# Check that the table is in the upper left corner
# The first_row is not working.
class Upper_Left_Corner(Test):

    def __init__(self, ws):
        Test.__init__(self, ws, "upper_left_corner")

        # Share that col_start and row_start are 0, default     
        self.first_row_idx = 0
        self.first_col_idx = 0

    def validate(self):

        
        mask = self.ws.notna().to_numpy()



        # Detect whether the headers were defined
        no_headers = self.ws.columns == "Unnamed"

        # If none of the headers are defined, search in the table for the them
        if no_headers.all():
            # Check each col for NA values, extract idx of first non-NA in each
            col_starts = np.argmax(mask, axis = 0)
            col_starts = np.where(
                np.sum(mask, axis = 0) == 0,
                len(col_starts),
                col_starts)

        # Else one or more of the headers were defined
        else:
            # So the first col is above the top row of the table
            col_starts = np.array([-1])


        # Check each row for NA values, extract column name of first non-NA
        row_starts = np.argmax(mask, axis = 1)
        row_starts = np.where(
            np.sum(mask, axis = 1) == 0,
            len(row_starts),
            row_starts)

        # Set the start of the grid
        # The first nonempty row is where first column starts
        first_row = int(col_starts.min())

        first_col = int(row_starts.min())

        
        # Check for blank space before the first column or row
        if first_col > 0 or first_row >= 0:
            
            # Where exactly it failed
            self.issues["first_col_idx"] = first_col
            self.issues["first_row_idx"] = first_row + 1

            # Save results to share with other tests
            self.first_col_idx = first_col
            self.first_row_idx = first_row + 1
            # Mark that there is a displaced table
            self.displaced = True
            # Save an effective table
            self.effective_ws = self.ws.iloc[first_row:, first_col:]
            # First row is the column names
            cols = self.effective_ws.iloc[0]
            self.effective_ws.columns = cols
            self.effective_ws = self.effective_ws[1:].reset_index(drop=True)
        else:
            self.displaced = False
            self.effective_ws = self.ws


        Test.validate(self,
            "Table is not in the upper left corner",
            "Table is in the upper left corner"
        )

# Detect if there are multiple tables
class Multi_Table(Test, Has_Dependency):

    def __init__(self, ws):
        Test.__init__(self, ws, "multi_table")
        Has_Dependency.__init__(self, Upper_Left_Corner)

    def validate(self, upper_left_corner):


        # Get shifted worksheet if relevant
        ws = upper_left_corner.effective_ws

        # Dimensions
        n_row, n_col = ws.shape

        # Array = 0 where value exists, else 1
        mask = ws.isna().to_numpy()
        # Obtain offset in case table not in upper left corner
        col_offset = upper_left_corner.first_col_idx
        row_offset = upper_left_corner.first_row_idx
        


        # First value in each row
        left_edge = np.argmin(mask, axis = 1)
        # left_edge = 0 if either all True or all False
        # test if all empty (True)
        empty_rows = np.all(mask, axis = 1)
        # Set left_edge to n_col if empty_row
        left_edge = np.where(empty_rows == 1, n_col, left_edge)
        # First value in each col
        top_edge = np.argmin(mask, axis = 0)
        # top_edge = 0 if either all True or all False
        # test if all empty (True)
        empty_cols = np.all(mask, axis = 0)
        
        # If there were any
        if np.any(empty_cols):
            # Could be multiple tables

            # Get locations of empty columns
            empty_cols_idx = np.argwhere(empty_cols)
            # Add an issue
            self.issues["empty columns"] = (empty_cols_idx + col_offset).tolist()
        
        # Set top_edge to n_row if empty_col
        top_edge = np.where(empty_cols == 1, n_row, top_edge)
        
        # Look for lower right corner
        # For candidate corner (i,j) in grid
        # If left_edge > j and also top_edge > i
        # We have a bounding box with
        # Upper right (0,0) and lower right (i,j)

        # Compute mesh grid of coordinates
        ii, jj = np.mgrid[0:n_row, 0:n_col]

        # Create zeros grid where empty cells left of left_edge filled with ones
        left_grid = np.where(left_edge[:,np.newaxis] > jj, 1, 0 )

        # Create zeros grid where empty cells above top_edge filled with ones
        top_grid = np.where(top_edge[np.newaxis,:] > ii, 1, 0)

        # If left padded grid intersects top padded grid
        # Then there exists a bounding box around subset of data
        box = np.logical_and(left_grid, top_grid)

        # If there is a bounding box smaller than the whole table
        if np.any(box):
            # Get location of lower right corner candidates
            lower_right_corners = np.argwhere(box)

            # Get location of corner closest to 0,0
            best_corner = lower_right_corners[0,:]

            # Add offset back in
            best_corner[0] = best_corner[0] + row_offset
            best_corner[1] = best_corner[1] + col_offset

            # Save the result as an issue
            self.issues["multi_table_corner"] = best_corner.tolist()

            # Feed forward property, prevents tests that would fail under multitable from running
            self.multi_table = True
        else:
            self.multi_table = False

        Test.validate(self,
            "Found bounding box around table smaller than full dimensions.  Sheet may have multiple tables.",
            "Sheet does not have multiple tables")


        
class Header(Test, Has_Dependency):

    def __init__(self, ws, name):

        # Initialize the test named headers
        Test.__init__(self, ws, name)

        # This test depends on the grid edges test
        Has_Dependency.__init__(self, Upper_Left_Corner, Multi_Table)

    def get_headers(self, upper_left_corner):
        # Check that the grid edges test has been run
        self.check_input(upper_left_corner)

        return upper_left_corner.effective_ws.columns
    
        


    


class Header_Duplicates(Header):

    def __init__(self, ws):

        # Initialize a headers test
        Header.__init__(self, ws, "header_duplicates")

    def validate(self, upper_left_corner, multi_table):
        if multi_table.multi_table:
            self.handle_multi_table()
        else:
            self.detect_duplicates(upper_left_corner)

    
    def detect_duplicates(self, upper_left_corner):
        # Get the headers
        headers = self.get_headers(upper_left_corner)

        # Extract duplicates
        duplicates = headers[headers.duplicated(keep = False)]

        # Check if there are duplicate headers
        if len(duplicates) > 0:

            # For each unique duplicate
            for dup in duplicates.unique():

                
                    # Flag the location as an issue
                    self.issues[dup] = f"Repeated {(duplicates == dup).sum()} times"

        Test.validate(self, "Duplicate headers found", "All headers unique")


class Header_ID(Header):

    def __init__(self, ws):
        Header.__init__(self, ws, "header_id")

    def validate(self, upper_left_corner, multi_table):
        if multi_table.multi_table:
            self.handle_multi_table()
        else:
            self.check_id(upper_left_corner)

    def check_id(self, upper_left_corner):
        headers = self.get_headers(upper_left_corner)
        # Iterate over headers
        if headers[0] == "ID":
            self.issues[(0, headers[0])] = "First header is ID"

        Test.validate(self, "First header is ID", "First header is not ID")

class Header_Length(Header):
        
        def __init__(self, ws):
            Header.__init__(self, ws, "header_length")

        def validate(self, upper_left_corner, multi_table):
            if multi_table.multi_table:
                self.handle_multi_table()
            else:
                self.check_length(upper_left_corner)

        def check_length(self, upper_left_corner):

            headers = self.get_headers(upper_left_corner)

            
            # Iterate over headers
            for idx, header in enumerate(headers):

                assert isinstance(header, str)

                # If bad length
                if len(header) <=3:
                    self.issues[(idx, header)] = "Header is less than 4 characters"
                
                if len(header) > 24:
                    self.issues[(idx, header)] = "Header is more than 24 characters"

            Test.validate(self, "Some headers have improper lengths", "All headers have acceptable lengths")
                


class Header_First_Char(Header):
        
        def __init__(self, ws):
            Header.__init__(self, ws, "header_first_char")

        def validate(self, upper_left_corner, multi_table):
            if multi_table.multi_table:
                self.handle_multi_table()
            else:
                self.check_first_char(upper_left_corner)

        def check_first_char(self, upper_left_corner):
            headers = self.get_headers(upper_left_corner)

            # Iterate over headers
            for idx, header in enumerate(headers):

                assert isinstance(header, str)

                # If bad length
                if header[0].isdigit():
                    self.issues[(idx, header)] = "Header starts with digit"

            Test.validate(self, "Some headers start with a digit", "No headers start with digits")
                


class Header_Space(Header):
        
        def __init__(self, ws):
            Header.__init__(self, ws, "header_space")

        def validate(self, upper_left_corner, multi_table):
            if multi_table.multi_table:
                self.handle_multi_table()
            else:
                self.check_space(upper_left_corner)

        def check_space(self, upper_left_corner):
            headers = self.get_headers(upper_left_corner)

            
            # Iterate over headers
            for idx, header in enumerate(headers):

                assert isinstance(header, str)

                # If bad length
                if " " in header:
                    self.issues[(idx, header)] = "Header has a space"
                

            Test.validate(self, "Some headers have spaces", "No headers have spaces")
                
class Header_Word_Separation(Header):

    def __init__(self, ws):
        Header.__init__(self, ws, "header_word_separation")

    def validate(self, upper_left_corner, multi_table):
        if multi_table.multi_table:
            self.handle_multi_table()
        else:
            self.check_underscore_dash(upper_left_corner)

    def check_underscore_dash(self, upper_left_corner):
        headers = self.get_headers(upper_left_corner)
        # Iterate over headers


        for idx, header in enumerate(headers):
            assert isinstance(header, str)

            # Detect captial letters in the middle of words
            camel_case = re.search(r'[a-z][A-Z]', header) is not None
            underscore =  "_" in header
            dash = "-" in header

        if sum([camel_case, underscore, dash]) > 1:
            self.issues[tuple(headers)] = "Headers mix camel case, underscores, and dashes"

            
        Test.validate(self, "Headers use mixture of camel case, underscores, and dashes", "Headers do not use a mixture of camel case, underscores and dashes")


class Header_Special_Characters(Header):
    def __init__(self, ws):
        Header.__init__(self, ws, "header_special_characters")

    def validate(self, upper_left_corner, multi_table):
        if multi_table.multi_table:
            self.handle_multi_table()
        else:
            self.check_special_characters(upper_left_corner)

    def check_special_characters(self, upper_left_corner):
        headers = self.get_headers(upper_left_corner)
        # Iterate over headers
        for idx, header in enumerate(headers):
            assert isinstance(header, str)
            if re.search(r"[!@#$%^&*()+=\[\]{};:'\"|\\,<>\?/]", header):
                self.issues[(idx, header)] = "Header contains special characters"

        Test.validate(self, "Headers contain special characters", "Headers do not contain special characters")

class Header_Date(Header):
    def __init__(self, ws):
        Header.__init__(self, ws, "header_date")

    def validate(self, upper_left_corner, multi_table):
        if multi_table.multi_table:
            self.handle_multi_table()
        else:
            self.check_date(upper_left_corner)

    def check_date(self, upper_left_corner):
        headers = self.get_headers(upper_left_corner)
        # Iterate over headers
        for idx, header in enumerate(headers):
            assert isinstance(header, str)
            if "date" in header.lower():
                self.issues[(idx, header)] = "Column may combine YYYY-MM-DD in one column instead of breaking up"

        Test.validate(self, "Column may combine YYYY-MM-DD in one column instead of breaking up", "Column does not combine YYYY-MM-DD in one column")


class Mixed_Datatypes(Test, Has_Dependency):

    def __init__(self, ws):

        # Initialize the test named mixed_datatypes
        Test.__init__(self, ws, "mixed_datatypes")

        # Needs to know if multiple tables are present
        Has_Dependency.__init__(self, Multi_Table, Upper_Left_Corner)

    def validate(self, multi_table, upper_left_corner):

        # If there are multiple tables on this sheet
        if multi_table.multi_table:
            # Pass by default
            Test.validate(self, None, "Multiple tables, cannot run test.  Pass by default.")    
        
        # Else there's a single table, run test
        else:
            # In case not in upper left corner, retrieve effective table
            ws = upper_left_corner.effective_ws

            n_rows, n_cols = ws.shape

            # For each column in the worksheet

            for col_idx, col_name in enumerate(ws.columns):

                types_found = dict()

                # Get the column data

                col_data = ws.iloc[:,col_idx]

                # Remove values that are missing to begin with

                col_data = col_data[col_data.notna()]

                # Attempt to convert the column to general numeric

                col_float = pd.to_numeric(col_data, errors='coerce')

                # Compute the number of rows that are possibly numeric

                n_numeric = col_float.notna().sum()

                # If there are any numeric rows

                if n_numeric > 0:

                    # Record the number of numeric entries

                    types_found["numeric"] = int(n_numeric)

                # # Attempt to convert the column to datetime in any iso format
                # col_datetime = pd.to_datetime(col_data, errors='coerce', format = "ISO8601")

                # # Compute number of datetimes
                # n_datetime = col_datetime.notna().sum()

                # # If there are any dates
                # if n_datetime > 0:

                #     types_found["datetime"] = int(n_datetime)


                # Compute number of missing data
                n_missing = col_data.isna().sum()

                # If it's not anything else, it must be text
                n_text = n_rows - n_numeric - n_missing

                # If there's any text
                if n_text > 0:
                    types_found["text"] = int(n_text)

                    # Loop to the next column
                    # Report each column containing multiple types
                    # Loop over the columns

                # If there are multiple types
                if len(types_found.keys()) > 1:

                    # Add to issues
                    self.issues[col_name] = types_found

                Test.validate(self,
                    "Table contains columns with mixed data types",
                    "Table columns each contains one datatype"
                )

class Aggregate_Row(Test, Has_Dependency):

    def __init__(self, ws):
        Test.__init__(self, ws, "aggregate_row")
        Has_Dependency.__init__(self, Multi_Table, Upper_Left_Corner)

    def validate(self, multi_table, upper_left_corner):
        if multi_table.multi_table:
            self.handle_multi_table()
        else:
            self.check_aggregate_row(upper_left_corner)

    def check_aggregate_row(self, upper_left_corner):
        ws = upper_left_corner.effective_ws

        # Get the last row
        last_row = ws.iloc[-1]

        # Search for the words aggregating words
        # List of indices of cells that contain aggregate words
        agg_word_idx = []
        # Regex to search: "total", "sum", "average", "count", "min", "max" 
        regex =r"(total|sum|average|count|min|max)"
        # Iterate over the cells in the last row
        for idx, cell_contents in enumerate(last_row):
            # If the cell contents is a string
            if isinstance(cell_contents, str):
                # If the cell contents contains an aggregate word
                if re.search(regex, cell_contents, flags = re.IGNORECASE):
                    # Note its index
                    agg_word_idx.append(idx)
        
        # Assume no aggregate words
        is_aggregate_row = False
        # For each aggregate word index
        for idx in agg_word_idx:
            # If the cell immediately to the right is a number
            if isinstance(last_row.iloc[idx + 1], (int, float)):
                # This is probably a true aggregate row
                is_aggregate_row = True
        
        # If it is an aggregate row
        if is_aggregate_row:
            # Add issues for each aggregate word
            for idx in agg_word_idx:
                row_idx = ws.index[-1]
                col = ws.columns[agg_word_idx[idx]]
                col_idx = agg_word_idx[idx]
                self.issues[row_idx, (col_idx, col)] = (last_row[idx], last_row[idx + 1])
        
        Test.validate(self, "Last row contains aggregate words", "Last row does not contain aggregate words")
        
# Look for special characters in the cells
class Special_Char(Test, By_Cell):

    def __init__(self, ws):
        # Initialize the test named special_characters
        Test.__init__(self, ws, "special_characters")

    def validate(self):
        # Use the default validate method for By_Cell
        By_Cell.validate(self,
            "Special characters found",
            "No special characters found")
    
    # Required by By_Cell
    # Does the cell contain special characters?
    def not_valid(self, cell):

        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # Check for special characters
        # This regex matches any character in the quotes
        # Bad_chars is a list of all the special characters found
        bad_chars = re.findall(r'[!@#$%&*(){}|<>/]', cell)

        # If there are no special characters, return False
        if len(bad_chars) == 0:
            return False
        else:
            # The cell had special characters
            return True

# Search for leading or trailing white space
class Untrimmed_White_Space(Test, By_Cell):
    def __init__(self, ws):
        Test.__init__(self, ws, "white_space")

    def validate(self):

        # Use the default validate method for By_Cell
        By_Cell.validate(self,
            "Leading or trailing white space found",
            "No leading or trailing hwhite space found")
    
    # Required by By_Cell
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # This regex matches any white space at the beginning or end of the string
        # Bad_space is a list of all the leading or trailing white space found
        bad_space = re.findall(r'^\s|\s$', cell)

        # If there is no leading or trailing white space, return False
        if len(bad_space) == 0:
            return False
        else:
            # The cell had leading or trailing white space
            return True

class Newlines_Tabs(Test, By_Cell):
    def __init__(self, ws):
        Test.__init__(self, ws, "newlines_tabs")

    def validate(self):

        # Use default validate method for By_Cell
        By_Cell.validate(self,
            "Newlines, tabs, or vertical tabs found.",
            "No newlines, tabs, or vertical tabs found.")
        
    # Required by By_Cell
    def not_valid(self, cell):

        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # This regex matches newlines, tabs, and vertical tabs anywhere
        # bad_space is a list of all the matches
        bad_space = re.findall(r'[\t\n\v]', cell)

        # If there are no newlines... etc
        if len(bad_space)==0:
            return False
        else:
            # The cell had newlines
            return True

# Needs to be redone as an analytic instead of a test

# class Quotes(Test, By_Cell):
#     def __init__(self, ws):
#         Test.__init__(self, ws, "quotes")

#     def validate(self):

#         # Use default validate method for By_Cell
#         By_Cell.validate(self,
#             "Quotes or double quotes found.",
#             "No quotes or double quotes found.")
        
#     # Required by By_Cell
#     def not_valid(self, cell):

#         # Check if the cell is a string
#         if not isinstance(cell, str):
#             # This test does not apply
#             return False
        
#         # This regex matches single and double quotes around comma or a tab
#         # bad_space is a list of all the matches
#         delimiters_in_cells = re.findall(r'[\'\"][,\t][\'\"]')

#         # If there are no newlines... etc
#         if len(delimiters_in_cells)==0:
#             return False
#         else:
#             # The cell had newlines
#             return True

class Missing_Value_Text(Test, By_Cell):
    def __init__(self, ws):
        Test.__init__(self, ws, "missing_value_text")

    def validate(self):

        # Use default validate method for By_Cell
        By_Cell.validate(self,
            "Text denoting missing values found.",
            "No text denoting missing values found.")
        
    # Required by By_Cell
    def not_valid(self, cell):

        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # This regex matches newlines, tabs, and vertical tabs anywhere
        # bad is a list of all the matches
        bad = re.findall(r'^(?:(no data)|(nd)|(missing)|(missing data)|(na)|(null)|(\-)|(\.)|(\s+)|(+_))$', cell, flags = re.IGNORECASE)

        # If there are matches
        if len(bad)==0:
            return False
        else:
            # The cell had missing data text
            return True

# Search for cells containing only question marks
class Question_Mark_Only(Test, By_Cell):
    def __init__(self, ws):
        Test.__init__(self, ws, "question_mark_only")

    def validate(self):
        By_Cell.validate(self, "Cells with just a question mark found", "No question mark cells found")
    
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # If the cell is exactly a question mark.
        return cell == "?"
            
class White_Space_Only(Test, By_Cell):
    def __init__(self, ws):
        Test.__init__(self, ws, "white_space_only")

    def validate(self):
        By_Cell.validate(self, "Cells with white space only found", "No cells with white space only found")
    
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        bad_space = re.findall(r'^\s+$', cell)
        if len(bad_space) == 0:
            return False
        else:
            return True

class Number_Space(Test, By_Cell):
    def __init__(self, ws):
        Test.__init__(self, ws, "white_space_only")

    def validate(self):
        By_Cell.validate(self, "Cells with only spaces and numbers found", "No cells with only spaces and numbers found")
    
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # Matches cells with at least one number and at least one spacer
        # but no other character type
        bad_match = re.findall(r'^(?=.*\d)(?=.*\s)[\d\s]+$', cell)
        if len(bad_match) == 0:
            return False
        else:
            return True



class Dates(Test, By_Cell):
    def __init__(self, ws):
        Test.__init__(self, ws, "dates")

    def validate(self):

        # Coerce all cells to datetimes, 'coerce' will set invalid dates to na
        # 'mixed' will try to parse the dates in any format
        parsed_dates = self.ws.apply(pd.to_datetime, errors='coerce', format = "mixed")

        # Check the ratio of valid dates to total dates
        # This will be a series with the column names as the index
        # and the ratio of valid dates as the values
        ratio_dates = parsed_dates.apply(lambda s: s.notna().mean())

        # Select the columns with more than 80% valid dates
        # Retrieve these columns names
        date_cols = ratio_dates[ratio_dates > 0.8].index
        
        # Select the columns from the full dataframe with dates
        date_cols = self.ws[date_cols]

        # Run the validation routine on just these columns
        By_Cell.validate(self, "Cells with non-ISO dates found", "Either no dates or all dates are ISO", date_cols)

    
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False

        # Check if the dates are in ISO format
        # This regex matches any date in the format YYYY-MM-DD
        cell_date = re.fullmatch(r'^\d{4}-?\d{2}-?\d{2}$', cell)
        if cell_date == None:
            return False
        else:
            return True


class Scientific_Notation(Test, By_Cell):
    # This test fails for "small" exponents like 1e-3
    # Because pandas reads these as floats automatically
    def __init__(self, ws):
        Test.__init__(self, ws, "scientific_notation")

    def validate(self):
        By_Cell.validate(self, "Cells with scientific notation found", "No cells with scientific notation found")
    
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # This regex matches any number in scientific notation
        sci_notation = re.fullmatch(r'^[+-]?\d+(?:\.\d+)?[eE][+-]?\d+$', cell.strip())

        # If the cell is not in scientific notation, return False
        if sci_notation == None:
            return False
        else:
            # The cell is in scientific notation
            return True




class Units(Test, By_Cell, Has_Dependency):
    # This test looks for units in the cells
    def __init__(self, ws):
        Test.__init__(self, ws, "units")
        Has_Dependency.__init__(self, Upper_Left_Corner)

    def validate(self, upper_left_corner):
        # Check that the grid edges test has been run
        self.check_input(upper_left_corner)

        By_Cell.validate(self, "Cells with units found", "No cells with units found", upper_left_corner.effective_ws)
    
    def not_valid(self, cell):
        
        # Check if the cell is a stringc
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # This regex matches any number with a unit suffix
        units = r'^[+-]?[\d\.]+(nm)|(um)|(mm)|(cm)|(m)|(km)|(in)|(ft)|(yd)|(mi)|(lb)|(oz)|(g)|(kg)|(mg)|(ug)|(ng)|(pg)|(ml)|(l)|(cl)|(ul)|(mmHg)|(inHg)|(psi)|(bar)|(atm)|(torr)|(s)|(ns)|(us)|(ms)|(hr)|(min)|(days)|(weeks)|(years)|(ml)|(dl)|(l)|(liters)|(gal)|(dbl)$'

        matches = re.fullmatch(units, cell.strip())

        # If the cell has no units, return False
        if units == None:
            return False
        else:
            # The cell has units
            return True




# If run as a script
if __name__ == "__main__":

    # Create a test suite from the demo notebook
    suite = Test_Suite("demo.xlsx")
    
    # Run all the tests and outputs
    suite.run()
    suite.report()
    suite.save(format = "json")
    suite.save(format = "csv")