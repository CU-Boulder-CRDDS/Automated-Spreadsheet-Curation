import pandas as pd # Reading excel files
from abc import ABC, abstractmethod # Creating abstract classes
import re # Regular expressions to test cell contents
import json # Saving results to json
from datetime import datetime as dt # Add timestamps to the results

class Test_Suite():
    __slots__ = ["wb_path", "wb", "to_run", "results"]

    # results is a dict of sheet names, each containing a dict of test names and their results
    # to_run is a list of test classes to run on each sheet
    # wb_path is the path to the workbook
    # wb is a dictionary of pandas dataframes, each representing a sheet in the workbook

    def __init__(self, wb_path, to_run = None):
        self.wb_path = wb_path

        # Assumes the input is excel (many versions supported)
        # Will need to add checks for other file types
        self.wb = pd.read_excel(
            io = wb_path, # The file path
            dtype = "str", # Parse all cells as strings, leaves the parsing to the tests
            sheet_name = None, # read all sheets returning a dict
            header = 0, # use first row as header, find later with grid_edges
        )

        # If no tests are provided, run all tests
        if to_run == None:
            self.to_run = [Grid_Edges, Header, Special_Char, Untrimmed_White_Space, Question_Marks, White_Space_Only, Dates]
        else:
            self.to_run = to_run

        # Initialize the results dict,
        # levels:  [sheet_name] -> [test_name] -> [test_object]
        self.results = dict()




    def run(self):
        # Run tests on the file as a whole

        # Initialize the results dict for the file
        self.results["file"] = dict()

        # Initialize the filename test
        filename_test = Filename(self.wb, self.wb_path)

        # Validate the filename test
        filename_test.validate()

        # Update the results dict with the filename test
        self.results["file"]["filename"] = filename_test


        # For each sheet in the workbook
        for sheet in self.wb.keys():
            print(f"Running tests on {sheet}")

            # Get the sheet as a dataframe
            ws = self.wb[sheet]

            # First check if the sheet is empty
            # Initialize the empty test
            empty_test = Empty(ws, self.wb_path)

            # Validate the empty test
            empty_test.validate()

            # If the test failed
            if not empty_test.status:

                # the sheet is empty, skip it
                print(f"Sheet {sheet} is empty. Skipping.")
                continue


            # Initialize the results dict for this sheet
            self.results[sheet] = dict()

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

            # Print the sheet name
            print(f"Sheet: {sheet}")

            # For each tests applied to this sheet
            # trimmed_results is a dict of test names and their results
            for test_name, trimmed_result in tests.items():
                
                # Print the report for that test
                print(f"Test: {test_name}")
                print(f"Message: {trimmed_result['message']}")
                print(f"Issues: {trimmed_result['issues']}")

                # Add some space
                print("\n")

            # Add extra extra space between sheets
            print("\n\n")

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
    def save(self, format = "json"):

        # If the requested format is json
        if format == "json":

            # Open a json file for writing
            with open(
                f"results{dt.now().strftime("%Y-%m-%d_%H-%M-%S")}.json", "w") as f:

                # Dump the trimmed results to the file
                json.dump(self.trimmed_results(stringify=True), f)

        # If the requested format is csv
        elif format == "csv":

            # Initialize a dataframe to contain the results
            df = pd.DataFrame(
                data = {
                    "path":[],
                    "sheet":[],
                    "test_name":[],
                    "message":[],
                    "issues":[]
                }
            )
            
            # For each sheet and its associated tests
            for sheet, tests in self.trimmed_results().items():

                # Make a dataframe of the test results in long form
                df_sheet = pd.DataFrame(
                    data = {
                        "path":[self.wb_path for _ in range(len(tests))],
                        "sheet":[sheet for _ in range(len(tests))],
                        "test_name":[test_name for test_name in tests.keys()],
                        "message":[test["message"] for test in tests.values()],
                        "issues":[test["issues"] for test in tests.values()]
                    }
                )
                
                # Append the test results to the dataframe
                df = pd.concat([df, df_sheet], ignore_index=True)

            # Save the dataframe to a csv file.
            df.to_csv(
                f"results{dt.now().strftime("%Y-%m-%d_%H-%M-%S")}.csv", mode = "w", index = False, header = True)

        # If the requested format is neither json nor csv
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


# Abstract class for tests that produce results used by other tests
# This class is used to create a feed forward structure
class Feed_Forward(ABC):

    def __init__(self):

        # The feed_forward field is added
        self.feed_forward = dict()

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

class Empty(Test):
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

class Filename(Test):
    def __init__(self, ws, wb_path):
        # Save the workbook path
        self.wb_path = wb_path

        # Initialize the test named filename
        Test.__init__(self, ws, "filename")

    def validate(self):

        # Check if the filename is too long
        if len(self.wb_path) > 200:

            # The test failed
            self.status = False

            # Set the message
            self.message = "Filename is too long"

            # Save the offending filename
            self.issues = {"filename": self.wb_path}

        else:
            # The test passed
            self.status = True

            # Set the message
            self.message = "Filename is not too long"

# Check that the table is in the upper left corner
# The first_row is not working.
class Grid_Edges(Test, Feed_Forward):

    def __init__(self, ws):
        Test.__init__(self, ws, "grid_edges")
        Feed_Forward.__init__(self)

        # Share that col_start and row_start are 0, default     
        self.feed_forward["first_row_idx"] = 0
        self.feed_forward["first_col_idx"] = 0

    def validate(self):
        
        # Check if the headers were empty when the sheet was imported
        col_unnamed = self.ws.columns.str.contains("Unnamed")

        # Check each col for NA values, extract idx of first non-NA in each
        # Add one to each column name assigned to "Unnamed"
        col_starts = col_unnamed + self.ws.notna().apply(
            lambda x : x.idxmax(), axis = 0)

        # Check each row for NA values, extract column name of first non-NA
        row_starts = self.ws.notna().apply(lambda x : x.idxmax(), axis = 1)

        # Convert the column names into integers
        row_starts = row_starts.map(lambda col: self.ws.columns.get_loc(col))

        # Set the start of the grid
        # The first nonempty row is where first column starts
        first_row = col_starts.min()

        # The first nonempty column is where first row starts
        first_col = row_starts.min()

        # Check for blank space before the first column or row
        if first_col != 0 or first_row != 0:
            
            # Where exactly it failed
            self.issues["first_col_idx"] = first_col
            self.issues["first_row_idx"] = first_row

            # Save results to share with other tests
            self.feed_forward["first_col_idx"] = first_col
            self.feed_forward["first_row_idx"] = first_row

        Test.validate(self,
            "Table is not in the upper left corner",
            "Table is in the upper left corner"
        )


class Header(Test, Has_Dependency):

    def __init__(self, ws):

        # Initialize the test named headers
        Test.__init__(self, ws, "headers")

        # This test depends on the grid edges test
        Has_Dependency.__init__(self, Grid_Edges)

            

    def validate(self, grid_edges):

        # Check that the grid edges test has been run
        self.check_input(grid_edges)

        # Get the start of the grid from the grid edges test
        first_row_idx = grid_edges.feed_forward["first_row_idx"]
        first_col_idx = grid_edges.feed_forward["first_col_idx"]

        # If the table is not positioned correctly
        if first_row_idx != 0 or first_col_idx != 0:
            # Select the header row
            # -1 because pandas makes the first row headers
            headers = self.ws.iloc[first_row_idx - 1, first_col_idx:]

        else: 
            headers = self.ws.columns

        # For each header and its index
        for idx, header in enumerate(headers):

            # If the header is valid
            if self.is_valid(header):

                # Go on to the next one
                continue
            else:

                # Set issue to header_location: invalid header
                self.issues[idx + first_col_idx] = header

        # Finish off setting messages and status
        Test.validate(self,
            "Headers are not acceptable",
            "Headers meet expectations"
        )

    # Check if the header is valid
    def is_valid(self, header):

        # Check if the header is a string
        # Other types, including NA, are not allowed.
        validity = isinstance(header, str)

        # Length greater than or equal to 3 and less than or equal to 12
        validity = validity and (len(header) >= 3 and len(header) <= 12)

        # First character is not a digit
        validity = validity and not header[0].isdigit()

        # No spaces in header
        validity = validity and not (" " in header)

        # Return result
        return validity

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

# Search for cells containing only question marks
class Question_Marks(Test, By_Cell):
    def __init__(self, ws):
        Test.__init__(self, ws, "question_marks")

    def validate(self):
        By_Cell.validate(self, "Question marks found", "No question marks found")
    
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
        # Retrieve the these columns names
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

# If run as a script
if __name__ == "__main__":

    # Create a test suite from the demo notebook
    suite = Test_Suite("demo.xlsx")
    
    # Run all the tests and outputs
    suite.run()
    suite.report()
    suite.save(format = "json")
    suite.save(format = "csv")