import pandas as pd
import numpy as np
import pprint
from abc import ABC, abstractmethod
import re
import json
from datetime import datetime as dt
from copy import deepcopy
import os
import inspect


def _validate_type_error(test_name, arg_name, expected, actual):
    expected_name = (
        expected.__name__
        if isinstance(expected, type)
        else " or ".join(t.__name__ for t in expected)
    )
    actual_name = type(actual).__name__
    raise ValueError(
        f"{test_name}: `{arg_name}` must be {expected_name}, got {actual_name}."
    )


def _validate_kwarg_type(test_name, arg_name, value, expected):
    if not isinstance(value, expected):
        _validate_type_error(test_name, arg_name, expected, value)


def _validate_optional_str_list(test_name, arg_name, value):
    if value is None:
        return []
    if not isinstance(value, list):
        _validate_type_error(test_name, arg_name, list, value)
    bad = [v for v in value if not isinstance(v, str)]
    if bad:
        raise ValueError(
            f"{test_name}: `{arg_name}` must be a list of strings. "
            f"Invalid values: {bad}"
        )
    return value


def _validate_regex_kwarg(test_name, arg_name, pattern):
    _validate_kwarg_type(test_name, arg_name, pattern, str)
    try:
        re.compile(pattern)
    except re.error as e:
        raise ValueError(
            f"{test_name}: `{arg_name}` must be a valid regex pattern. Error: {e}"
        ) from e
    return pattern


def _discover_tests(test_level = "all", to_run = None, to_skip = None):
    """
    Find all concrete (non-abstract) Test subclasses whose __init__ takes only
    `ws` as a required positional arg.  These are the per-sheet tests that
    Test_Suite queues automatically.
    """

    # Save either to_run or to_skip as a set to filter the tests
    if to_run is not None:
        to_filter = set(to_run)
        is_included = True
    elif to_skip is not None:
        to_filter = set(to_skip)
        is_included = False
    else:
        # Else we are not including the empty set
        to_filter = set()
        is_included = False



    # List of all available tests
    all_tests = dict()

    # List of classes to assess
    to_visit = list(Test.__subclasses__())

    # Classes we've already assessed
    # Prevents infinite loops from diamond dependency graphs
    visited = set()

    # While there are classes to assess
    while to_visit:

        # Get the first class
        cls = to_visit.pop(0)

        # If the class has already been assessed, skip it
        if cls in visited:
            # Skip it
            continue

        # Add the class to the list of visited classes
        visited.add(cls)

        # Add the subclasses of the class to the list of classes to assess
        to_visit.extend(cls.__subclasses__())

        # If the class is abstract
        if inspect.isabstract(cls):
            # Skip it
            continue
        
        # Else the class is not abstract
        else:

            # Add the test_name to the list of all tests
            all_tests[cls.__name__.lower()] = cls
    
    

    # Verify that to_filter contains valid test names
    for test_name in to_filter:
        # If the test name is not in the list of all tests
        if test_name not in all_tests:
            # Raise an error
            raise ValueError(f"Invalid test name: {test_name}")


    # Initialize the result list of tests
    result = {}

    # For each test in the list of all tests
    for test_name,test_cls in all_tests.items():

        # If test_level
        if test_level == "all" or test_level in test_name:

            # Apply the filter
            if test_name in to_filter and is_included:
                # Add the class to the list of tests
                result[test_name] = test_cls
            elif test_name not in to_filter and not is_included:
                # Add the class to the list of tests
                result[test_name] = test_cls
            else:
                # Skip the class
                continue
            
    # Return the dictionary of tests
    return result


class Test_Suite():

    # results is a dict of sheet names, each containing a dict of test names and their results
    # to_run is a list of test classes to run on each sheet
    # wb_path is the path to the workbook
    # wb is a dictionary of pandas dataframes, each representing a sheet in the workbook

    def __init__(self, wb_path, to_run=None, to_skip=None, config_path=None):
        self.wb_path = wb_path

        # --- load config from JSON (optional) ---
        self.config = {}
        if config_path is not None:
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            except FileNotFoundError:
                print(f"Warning: Config file '{config_path}' not found. Using defaults.")
            except json.JSONDecodeError as e:
                print(f"Warning: Config file '{config_path}' has invalid JSON: {e}. Using defaults.")


        # --- validate to_run / to_skip mutual exclusivity ---
        if to_run is not None and to_skip is not None:
            raise ValueError("Cannot specify both to_run and to_skip. Use one or neither.")
        
        to_filter = to_run if to_run else to_skip
        if to_filter:
            if not isinstance(to_filter, list):
                raise ValueError("to_run or to_skip must be a list (of test names).")
            if not all(isinstance(test_name, str) for test_name in to_filter):
                raise ValueError("to_run or to_skip lists must be of test names (strings).")


        # Discover all the tests to run
        all_tests = _discover_tests(test_level = "all", to_run = to_run, to_skip = to_skip)

        # Verify that config only contains valid test names
        for test_name in self.config.keys():
            if test_name not in all_tests:
                raise ValueError(f"Invalid test name in config: {test_name}. Using specified tests from to_run or to_skip, allowed test names are: {list(all_tests.keys())}.")


        # Initialize the results dict,
        # levels:  [sheet_name] -> [test_name] -> [test_object]
        self.results = dict()
        # Initialize the results dict for the file
        self.results["file"] = dict()



        # First the file level tests
        self.file_tests = {t_name: self._create_test(t_cls) for t_name, t_cls in all_tests.items() if t_name.startswith("file")}


        # --- File encoding test ---
        # Must happen before loading pandas dataframe

        # First check if the file is UTF-8 encoded
        if "file_encoding" in self.file_tests:
            encoding_test = self.file_tests["file_encoding"]
            encoding_test.validate(self.wb_path)
            encoding = encoding_test.valid_encoding
            self.results["file"]["file_encoding"] = encoding_test

            # Remove the file encoding test from remaining tests to run
            self.file_tests.pop("file_encoding")
        else:
            # Skip the file encoding test, default to UTF-8
            encoding = "utf-8"
            pass


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
                encoding = encoding, # Read the file
                encoding_errors = "replace", # Replace invalid characters with ?
            )
            # Make the single sheet into a singleton dict to match excel format
            self.wb = {os.path.splitext(os.path.basename(wb_path))[0]: self.wb}
        else:
            raise ValueError(f"File {wb_path} has an invalid file extension: {file_extension}. Use '.xlsx', or '.csv'.  If you want to use '.xls', use conda to install xlrd and change the if statement above to allow xls.")


        # --- Check input arguments from config to tests
        
        
        # Initialize the sheet tests to check input arguments
        self.sheet_tests = {t_name: self._create_test(t_cls) for t_name, t_cls in all_tests.items() if t_name.startswith("sheet")}

        # Initialize the header tests to check input arguments
        self.header_tests = {t_name: self._create_test(t_cls) for t_name, t_cls in all_tests.items() if t_name.startswith("header")}

        # Initialize the cell tests to check input arguments
        self.cell_tests = {t_name: self._create_test(t_cls) for t_name, t_cls in all_tests.items() if t_name.startswith("cell")}


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

        if isinstance(self.wb, dict):
            # Get each dataframe in the dict
            self.wb = {name : set_headers(df) for name, df in self.wb.items()}
        else:
            self.wb = set_headers(self.wb)







    def _create_test(self, test_cls):
        """
        Instantiate a test, merging in any config overrides for its
        keyword-only parameters.  First creates with defaults to discover the
        test name, then re-creates with config overrides if any exist.
        """
        test_name = test_cls.__name__.lower()

        # Get the users config options for this test
        cfg = self.config.get(test_name, {})


        # Get the call signature of the tests
        sig = inspect.signature(test_cls.__init__)
        # The allowed options are the keyword-only parameters
        allowed = {
            name for name, p in sig.parameters.items()
            if p.kind == inspect.Parameter.KEYWORD_ONLY
        }
        # Check for any invalid options provided by the user
        bad_options = {k: v for k, v in cfg.items() if k not in allowed}
        if bad_options:
            raise ValueError(f"Invalid options for test {test_name}: {bad_options}. Allowed options are: {allowed}.  See the README documentation for more information.")
        # Get the valid options provided by the user (which might be none)
        overrides = {k: v for k, v in cfg.items() if k in allowed}
        
        # Create the test with the valid options provided by the user
        test = test_cls(**overrides)
        return test


    def _validate_tests(self, tests, positional_args, other_dependencies = None):
        f"""
        Runs a list of tests with specified args.

        Args:
            tests: A list of test objects to run.
            positional_args: A list of positional arguments for all the tests.
            other_dependencies: (Not implemented) dict of completed dependencies, dep_name : completed_dep_test, to pass to the tests that are not already in the list of tests.
        """
        import traceback

        completed_tests = dict()
        finalized_tests = dict()
        queue = list(tests.values())
        skip_count = 0

        # While there are tests to run
        while len(queue) > 0:
            try:
                # Check we haven't skipped every test in the queue consecutively.
                # If so, no remaining test is currently runnable.
                if skip_count >= len(queue):
                    remaining_test_names = [t.name for t in queue]
                    raise RuntimeError(
                        f"Circular or unmet dependencies among remaining tests: {remaining_test_names}"
                    )

                # Get the first test
                t_obj = queue.pop(0)

                # Check if the test has dependencies
                if isinstance(t_obj, Has_Dependency):
                    fulfilled = {
                        dep_name: dep_test
                        for dep_name, dep_test in completed_tests.items()
                        if type(dep_test) in t_obj.dependencies
                    }

                    if len(fulfilled) < len(t_obj.dependencies):
                        # Add the test back to the queue
                        queue.append(t_obj)
                        # Increment the skip count
                        skip_count += 1
                        continue
                else:
                    # Fulfilled is empty
                    fulfilled = dict()

                # Validate the test
                try:
                    t_obj.validate(*positional_args, **fulfilled)
                    # Track the completed test
                    completed_tests[t_obj.name] = t_obj
                    finalized_tests[t_obj.name] = t_obj
                    # Reset the skip count because we've completed a test
                    skip_count = 0
                except Exception as e:
                    user_kwonly_args = self.config.get(t_obj.name, {})
                    print(
                        "Warning: Runtime error while executing test "
                        f"'{t_obj.name}'.\n"
                        f"  positional_args={positional_args}\n"
                        f"  user_keyword_only_args={user_kwonly_args}\n"
                        f"  error_type={type(e).__name__}\n"
                        f"  error_message={e}\n"
                        f"  traceback:\n{traceback.format_exc()}"
                    )
                    # Runtime-error tests are treated as not completed.
                    t_obj.status = None
                    t_obj.message = (
                        f"Test did not complete due to runtime error: {type(e).__name__}: {e}"
                    )
                    finalized_tests[t_obj.name] = t_obj
                    # We made progress by removing this test from the queue.
                    skip_count = 0

            except RuntimeError:
                remaining_test_names = [t.name for t in queue]
                print(
                    "Warning: Circular or unmet dependencies detected. "
                    "The following tests remained in the queue and will not be completed: "
                    f"{remaining_test_names}"
                )
                # Mark all remaining tests as not completed and stop this loop.
                for t_obj in queue:
                    t_obj.status = None
                    t_obj.message = (
                        "Test did not complete due to circular or unmet dependencies."
                    )
                    finalized_tests[t_obj.name] = t_obj
                break

        # Return all tests (completed + not completed)
        return finalized_tests



    def run(self):
        # --- file-level tests ---
        print("Running file-level tests")
        self.results["file"].update(
            self._validate_tests(self.file_tests, [self.wb_path]))

        print("Running sheet, header, and cell tests")
        remaining_tests = self.sheet_tests | self.header_tests | self.cell_tests

        for sheet, ws in self.wb.items():
            fresh_tests = deepcopy(remaining_tests)
            

            self.results[sheet] = self._validate_tests(fresh_tests, [ws, sheet])

            
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


    def __init__(self):
        # Get the base class name (deepest/first concrete class in MRO)
        base_cls = type(self)
        # Save the base class name as the test name
        self.name = base_cls.__name__.lower()

        # Whether the test has been run
        self.is_run = False

        # True if the test passed, False if it failed
        self.status = None


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


    def handle_empty(self):
        Test.validate(
            self,
            None,
            "Cannot assess test if sheet is empty.  Pass by default")

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
        if hasattr(self, "dependencies"):
            self.dependencies += list(dependencies)
        else:
            self.dependencies = list(dependencies)

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



class File(Test, ABC):
    def __init__(self,):

        # Initialize the test
        Test.__init__(self,)

    def set_positional(self, wb_path):
        if not isinstance(wb_path, str):
            raise ValueError("wb_path must be a string")
        self.wb_path = wb_path


class File_Name(File, ABC):
    def __init__(self,):

        # Initialize the test
        File.__init__(self,)

    
    def set_positional(self, wb_path):
        File.set_positional(self, wb_path)

        # Save the filename
        self.filename = os.path.basename(self.wb_path)


        
class File_Name_Length(File_Name):
    def __init__(self, *, min_length=5, max_length=32):

        File_Name.__init__(self,)
        _validate_kwarg_type(self.name, "min_length", min_length, int)
        _validate_kwarg_type(self.name, "max_length", max_length, int)
        if min_length < 0 or max_length < 0:
            raise ValueError(f"{self.name}: `min_length` and `max_length` must be >= 0.")
        if min_length >= max_length:
            raise ValueError(
                f"{self.name}: `min_length` ({min_length}) must be less than "
                f"`max_length` ({max_length})."
            )
        self.min_length = min_length
        self.max_length = max_length

    def validate(self, wb_path):
        File_Name.set_positional(self, wb_path)

        # If length is greater than max_length.
        if len(self.filename) > self.max_length:
            # The test failed
            self.status = False
            # Add an issue "location = filename": "contents of filename"
            self.issues["filename"] = self.filename
            # Set message
            self.message = f"Filename is > {self.max_length} characters"
        elif len(self.filename) <= self.min_length:
            self.status = False
            self.issues["filename"] = self.filename
            self.message = f"Filename is < {self.min_length} characters"
        else:
            # The test passed
            self.status = True
            # Set message
            self.message = f"Filename is < {self.max_length} characters"


class File_Name_Whitespace(File_Name):
    def __init__(self,):
        # Initialize the test
        File_Name.__init__(self,)

    def validate(self, wb_path):
        File_Name.set_positional(self, wb_path)

        # detect whitespace in filename
        # search returns None if no match is found, else Match object
        if re.search(r"\s", self.filename) is not None:

            # The test failed
            self.status = False

            # Add an issue
            self.issues["filename"] = self.filename
        Test.validate(self, "Filename contains whitespace", "Filename does not contain whitespace")

class File_Name_Final(File_Name):
    def __init__(self,):
        # Initialize the test named filename_final
        File_Name.__init__(self,)

    def validate(self, wb_path):
        File_Name.set_positional(self, wb_path)

        # If filename contains 'final'
        if "final" in self.filename.lower():
            self.status = False

            self.issues["filename"] = self.filename

            self.message = "Filename contains 'final'"
        else:
            self.status = True
            self.message = "Filename does not contain 'final'"
    
class File_Name_Word_Separation(File_Name):
    def __init__(self,):
        # Initialize the test named filename_word_separation
        File_Name.__init__(self,)

    def validate(self, wb_path):
        File_Name.set_positional(self, wb_path)

        # Detect combo of camel case, underscore, and dash in filename
        # Expressions work even if there are whitespace in the filename
        # camel case: capital letters in the middle of words
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

class File_Name_Special_Characters(File_Name):
    def __init__(self, *,
                 special_char_pattern=r"[!@#$%^&*()+=\[\]{};:'\"|\\,<>\?/]"):
        File_Name.__init__(self,)
        self.special_char_pattern = _validate_regex_kwarg(
            self.name, "special_char_pattern", special_char_pattern
        )

    def validate(self, wb_path):
        File_Name.set_positional(self, wb_path)

        spec_char = re.search(self.special_char_pattern, self.filename)
        
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
        

class File_Encoding(File):
    def __init__(self, *, valid_encoding="utf-8"):
        
        File.__init__(self,)
        _validate_kwarg_type(self.name, "valid_encoding", valid_encoding, str)
        self.valid_encoding = valid_encoding
    
    def validate(self, wb_path):
        File.set_positional(self, wb_path)
        self.file_extension = os.path.splitext(self.wb_path)[1].lower()
        
        # Only check encoding for text-based files (CSV, not Excel)
        if self.file_extension == ".xlsx" or self.file_extension == ".xls":
            # Excel files are binary ZIP archives, encoding check doesn't apply
            self.status = True
            self.message = "Encoding check skipped for Excel files (binary format)"
            self.is_run = True
            return
        
        
        # For CSV and other text files, check encoding
        try:
            with open(self.wb_path, 'rb') as f:
                # Read the file content as a list of lines in bytes
                lines_bytes = f.readlines()
                # Try to decode each line as UTF-8
                for row_num, line in enumerate(lines_bytes):
                    text = line.decode(self.valid_encoding, errors='replace')

                    # Check to see if the official replacement character was used
                    if '\ufffd' in text:
                        self.issues[f"row {row_num}"] = text
                        
        except Exception as e:
            # Handle other exceptions (e.g., file not found)
            self.status = False
            self.issues["file"] = f"Error reading file: {str(e)}"
            self.message = f"Error checking file encoding: {str(e)}"
        
        # Pass issues and messages forward
        Test.validate(self, f"File encoding is not {self.valid_encoding}", f"File encoding is {self.valid_encoding}")




class Sheet(Test, Has_Dependency, ABC):
    def __init__(self,):
        Test.__init__(self)
        if not isinstance(self, Sheet_Empty):
            Has_Dependency.__init__(self, Sheet_Empty)
        else:
            self.dependencies = []


    def set_positional(self, ws, ws_name):
        if not isinstance(ws, pd.DataFrame):
            raise ValueError("ws must be a pandas DataFrame")
        if not isinstance(ws_name, str):
            raise ValueError("ws_name must be a string")
        self.ws = ws
        self.ws_name = ws_name

class Sheet_Empty(Sheet):
    def __init__(self,):
        
        Sheet.__init__(self,)
        self.empty = None


    def validate(self, ws, ws_name):
        self.set_positional(ws, ws_name)

        # Check if the worksheet is empty
        if self.ws.empty:
            # The test failed
            self.status = False

            # Set empty status
            self.empty = True

            # Set the message
            self.message = "Worksheet is empty"

            self.issues = {"sheet": self.ws_name}

        else:
            # The test passed
            self.status = True

            # Set empty status
            self.empty = False

            # Set the message
            self.message = "Worksheet is not empty"

class Sheet_Name(Sheet):
    def __init__(self, *,
                 special_char_pattern=r"[!@#$%^&*()+=\[\]{};:'\"|\\,<>\?/]"):
        
        Sheet.__init__(self,)
        self.special_char_pattern = _validate_regex_kwarg(
            self.name, "special_char_pattern", special_char_pattern
        )

    def validate(self, ws, ws_name, sheet_empty):
        self.set_positional(ws, ws_name)



        # detect special characters in filename
        spec_char = re.search(self.special_char_pattern, self.ws_name)
        message = "Worksheet name: "
        if spec_char:
            self.status = False

            self.issues["sheet name"] = spec_char.group(0)
            message += f"contains special characters ({spec_char.group(0)}) "
        
        # detect whitespace in sheet name
        if re.search(r"\s", self.ws_name):
            self.status = False

            self.issues["sheet name"] = self.ws_name
            message += "contains whitespace "

        
        if  self.status is None:
            # The test passed
            self.status = True

            # Set the message
            self.message = "Worksheet name does not contain special characters or whitespace"
        else:
            self.message = message

# Check that the table is in the upper left corner
# The first_row is not working.
class Sheet_Upper_Left_Corner(Sheet):

    def __init__(self,):
        Sheet.__init__(self,)

        # Share that col_start and row_start are 0, default     
        self.first_row_idx = 0
        self.first_col_idx = 0

    def validate(self, ws, ws_name, sheet_empty):
        self.set_positional(ws, ws_name)

        
        mask = self.ws.notna().to_numpy()

        if sheet_empty.empty:
            self.handle_empty()
            return



        # Detect whether the headers were defined
        no_headers = self.ws.columns == "Unnamed"

        # If none of the headers are defined, search in the table for them
        if no_headers.all():
            # Check each col for NA values, extract idx of first non-NA in each
            col_starts = np.argmax(mask, axis = 0)
            col_starts = np.where(
                np.sum(mask, axis = 0) == 0,
                len(col_starts),
                col_starts)

        # Else one or more of the headers were defined
        else:
            # So the first col is the top row of the table
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
        first_row = first_row + 1

        first_col = int(row_starts.min())

        
        # Check for blank space before the first column or row
        if first_col > 0 or first_row > 0:
            
            # Where exactly it failed
            self.issues["first_col_idx"] = first_col
            self.issues["first_row_idx"] = first_row

            # Save results to share with other tests
            self.first_col_idx = first_col
            self.first_row_idx = first_row
            # Mark that there is a displaced table
            self.displaced = True
            # Save an effective table
            self.effective_ws = self.ws.iloc[first_row:, first_col:]
            if first_row > 0:
                # First row is the column names
                cols = self.effective_ws.iloc[0]
                self.effective_ws.columns = cols
                self.effective_ws = self.effective_ws[1:].reset_index(drop=True)
            # Else the columns are already assigned correctly

        else:
            self.displaced = False
            self.effective_ws = self.ws


        Test.validate(self,
            "Table is not in the upper left corner",
            "Table is in the upper left corner"
        )

# Detect if there are multiple tables
class Sheet_Multi_Table(Sheet, Has_Dependency):

    def __init__(self,):
        Sheet.__init__(self,)
        Has_Dependency.__init__(self, Sheet_Upper_Left_Corner)

    def validate(self, ws, ws_name, sheet_empty, sheet_upper_left_corner):
        self.set_positional(ws, ws_name)

        if sheet_empty.empty:
            self.handle_empty()
            return

        # Get shifted worksheet if relevant
        ws = sheet_upper_left_corner.effective_ws

        # Dimensions
        n_row, n_col = ws.shape

        # Array = 0 where value exists, else 1
        mask = ws.isna().to_numpy()
        # Obtain offset in case table not in upper left corner
        col_offset = sheet_upper_left_corner.first_col_idx
        row_offset = sheet_upper_left_corner.first_row_idx
        


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
            self.issues["empty columns"] = (empty_cols_idx + col_offset).ravel().tolist()
        
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

    def __init__(self):
        Test.__init__(self)

        # This test depends on the grid edges test
        Has_Dependency.__init__(self, Sheet_Empty,Sheet_Upper_Left_Corner, Sheet_Multi_Table)

    def _handle_dependencies(self, sheet_empty, sheet_upper_left_corner, sheet_multi_table, fn):
        if sheet_empty.empty:
            self.handle_empty()
        elif sheet_multi_table.multi_table:
            self.handle_multi_table()
        else:
            self.headers = self.get_headers(sheet_upper_left_corner)
            fn(self.headers)


    def get_headers(self, sheet_upper_left_corner):
        # Check that the upper left corner test has been run
        self.check_input(sheet_upper_left_corner)

        return sheet_upper_left_corner.effective_ws.columns
    
        




class Header_Duplicates(Header):

    def __init__(self):

        # Initialize a headers test
        Header.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self.detect_duplicates)

    def detect_duplicates(self, headers):
        duplicates = headers[headers.duplicated(keep = False)]
        if len(duplicates) > 0:
            for dup in duplicates.unique():
                    # Flag the location as an issue
                    self.issues[dup] = f"Repeated {(duplicates == dup).sum()} times"

        Test.validate(self, "Duplicate headers found", "All headers unique")


class Header_ID(Header):

    def __init__(self):
        Header.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self.check_id)

    def check_id(self, headers):
        if headers[0] == "ID":
            self.issues[(0, headers[0])] = headers[0]

        Test.validate(self, "First header is ID", "First header is not ID")

class Header_Length(Header):
        
        def __init__(self, *, min_length=1, max_length=24):
            Header.__init__(self)
            _validate_kwarg_type(self.name, "min_length", min_length, int)
            _validate_kwarg_type(self.name, "max_length", max_length, int)
            if min_length < 0 or max_length < 0:
                raise ValueError(
                    f"{self.name}: `min_length` and `max_length` must be >= 0."
                )
            if min_length > max_length:
                raise ValueError(
                    f"{self.name}: `min_length` ({min_length}) must be <= "
                    f"`max_length` ({max_length})."
                )
            self.min_length = min_length
            self.max_length = max_length

        def validate(self, *ignore, **dependencies):
            self._handle_dependencies(**dependencies, fn=self.check_length)

        def check_length(self, headers):
            for idx, header in enumerate(headers):

                assert isinstance(header, str)

                # If bad length
                if len(header) < self.min_length or len(header) > self.max_length:
                    self.issues[(idx, header)] = header


            Test.validate(self, "Headers should be between 1 and 24 characters", "All headers have acceptable lengths")
                


class Header_First_Char(Header):
        
        def __init__(self):
            Header.__init__(self)

        def validate(self, *ignore, **dependencies):
            self._handle_dependencies(**dependencies, fn=self.check_first_char)

        def check_first_char(self, headers):
            for idx, header in enumerate(headers):

                assert isinstance(header, str)

                # If bad length
                if header[0].isdigit():
                    self.issues[(idx, header)] = header

            Test.validate(self, "Some headers start with a digit", "No headers start with digits")
                


class Header_Space(Header):
        
        def __init__(self):
            Header.__init__(self)

        def validate(self, *ignore, **dependencies):
            self._handle_dependencies(**dependencies, fn=self.check_space)

        def check_space(self, headers):
            for idx, header in enumerate(headers):

                assert isinstance(header, str)

                # If bad length
                if " " in header:
                    self.issues[(idx, header)] = header
                

            Test.validate(self, "Some headers have spaces", "No headers have spaces")
                
class Header_Word_Separation(Header):

    def __init__(self):
        Header.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self.check_underscore_dash)

    def check_underscore_dash(self, headers):
        for idx, header in enumerate(headers):
            assert isinstance(header, str)

            # Detect captial letters in the middle of words
            camel_case = re.search(r'[a-z][A-Z]', header) is not None
            underscore =  "_" in header
            dash = "-" in header

        if sum([camel_case, underscore, dash]) > 1:
            self.issues["all headers"] = headers

            
        Test.validate(self, "Headers use mixture of camel case, underscores, and dashes", "Headers do not use a mixture of camel case, underscores and dashes")


class Header_Special_Characters(Header):
    def __init__(self, *,
                 special_char_pattern=r"[!@#$%^&*()+=\[\]{};:'\"|\\,<>\?/]"):
        Header.__init__(self)
        self.special_char_pattern = _validate_regex_kwarg(
            self.name, "special_char_pattern", special_char_pattern
        )

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self.check_special_characters)

    def check_special_characters(self, headers):
        for idx, header in enumerate(headers):
            assert isinstance(header, str)
            if re.search(self.special_char_pattern, header):
                self.issues[(idx, header)] = header

        Test.validate(self, "Headers contain special characters", "Headers do not contain special characters")

class Header_Date(Header):
    def __init__(self, *, date_keywords=None):
        Header.__init__(self)
        if date_keywords is None:
            date_keywords = ["date", "datetime", "timestamp"]
        self.date_keywords = _validate_optional_str_list(
            self.name, "date_keywords", date_keywords
        )

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self.check_date)

    def check_date(self, headers):
        for idx, header in enumerate(headers):
            assert isinstance(header, str)
            if any(keyword in header.lower() for keyword in self.date_keywords):
                self.issues[(idx, header)] = header

        Test.validate(self, "Column may combine YYYY-MM-DD in one column instead of breaking up", "Column does not combine YYYY-MM-DD in one column")


class Header_Mixed_Datatypes(Test, Has_Dependency):

    def __init__(self):

        # Initialize the test named mixed_datatypes
        Test.__init__(self)

        # Needs to know if multiple tables are present
        Has_Dependency.__init__(self, Sheet_Multi_Table, Sheet_Upper_Left_Corner)

    def validate(self, ws, ws_name, **dependencies):
        self.ws = ws
        if dependencies["sheet_multi_table"].multi_table:
            self.handle_multi_table()
        else:
            self.check_mixed_datatypes()
        
    def check_mixed_datatypes(self):
        ws = self.ws
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

# Abstract class for tests that check each cell in the worksheet
class Cell(Test, Has_Dependency, ABC):

    def __init__(self):
        Test.__init__(self)
        Has_Dependency.__init__(
            self, Sheet_Upper_Left_Corner, Sheet_Multi_Table)

    # Generator over pandas dataframes
    # tuple of the cell coordinates and the cell contents
    # This is used to iterate over the cells in the worksheet
    def pandas_iter(self, df = None):

        # If no dataframe is provided, use the worksheet
        if df is None:
            df = self.effective_ws

        # For each row in the worksheet except the first row (headers)
        for row_idx, row in df.iterrows():

            # For each column
            for col_idx, col_name in enumerate(df.columns):
                cell = df.iloc[row_idx, col_idx]

                # Yield the cell coordinates and the cell contents
                # (row_idx, (col_idx, col_name)), contents)
                yield (row_idx, (col_idx, col_name)), cell

    
    def _handle_dependencies(self, sheet_upper_left_corner, sheet_multi_table, fn):
        if sheet_multi_table.multi_table:
            self.handle_multi_table()
        else:
            fn(sheet_upper_left_corner.effective_ws)

    # Cell by cell validation.  
    # Assumes that the not_valid method is implemented
    def validate(self, fail_message, pass_message, df = None):

        # If provided, df may be a subset of the whole worksheet
        # If not provided, use the effective worksheet
        if df is None:
            df = self.ws

        
        for cell in self.pandas_iter(df):

            # If the cell is not valid, as given by subclass
            if self.not_valid(cell[1]):

                # Save the cell coordinates and the cell contents
                self.issues[cell[0]] = cell[1]
        

        # Set the status messages
        Test.validate(self, fail_message, pass_message)


    # The not_valid method is an abstract method that must be implemented
    # by each test class that inherits from Cell
    # This method should return True if the cell is not valid, and False if it is valid
    # The validate method will call this method for each cell in the worksheet
    @abstractmethod
    def not_valid(self, cell):
        pass


class Cell_Aggregate_Row(Cell):

    def __init__(self, *, aggregate_words=None):
        Cell.__init__(self)

        if aggregate_words is None:
            aggregate_words = ["total", "sum", "average", "count", "min", "max"]
        self.aggregate_words = _validate_optional_str_list(
            self.name, "aggregate_words", aggregate_words
        )

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_aggregate_row)

    def _check_aggregate_row(self, df):

        words = self.aggregate_words
        if not words:
            regex = None
        else:
            regex = re.compile(
                r"(" + "|".join(re.escape(w) for w in words) + r")",
                re.IGNORECASE,
            )

        # Get the last row
        last_row = df.iloc[-1]

        # List of indices of cells that contain aggregate words
        agg_word_idx = []
        for idx, cell_contents in enumerate(last_row):
            if isinstance(cell_contents, str) and regex is not None:
                if regex.search(cell_contents):
                    agg_word_idx.append(idx)

        is_aggregate_row = False
        for idx in agg_word_idx:
            if idx + 1 < len(last_row) and isinstance(
                last_row.iloc[idx + 1], (int, float)
            ):
                is_aggregate_row = True
                break

        if is_aggregate_row:
            row_idx = df.index[-1]
            for idx in agg_word_idx:
                col_idx = idx
                col = df.columns[col_idx]
                self.issues[row_idx, (col_idx, col)] = (
                    last_row.iloc[idx],
                    last_row.iloc[idx + 1],
                )

        Test.validate(self, "Last row contains aggregate words", "Last row does not contain aggregate words")

    # Required by Cell to make this class not abstract
    def not_valid(self, cell):
        pass
        
# Look for special characters in the cells
class Cell_Special_Characters(Cell):

    def __init__(self, *,
                 default_pattern=r"[!@#$%^&*()+=\[\]{};:'\"|\\,<>\?/]",
                 url_pattern=r"^https?://",
                 free_text_pattern=r"[!@#$%&*(){}\[\]|<>/]",
                 url_columns=None,
                 free_text_columns=None,
                 skip_columns=None):
        Cell.__init__(self)
        
        self.default_pattern = _validate_regex_kwarg(
            self.name, "default_pattern", default_pattern
        )
        self.url_pattern = _validate_regex_kwarg(self.name, "url_pattern", url_pattern)
        self.free_text_pattern = _validate_regex_kwarg(
            self.name, "free_text_pattern", free_text_pattern
        )
        self.url_columns = _validate_optional_str_list(
            self.name, "url_columns", url_columns
        )
        self.free_text_columns = _validate_optional_str_list(
            self.name, "free_text_columns", free_text_columns
        )
        self.skip_columns = _validate_optional_str_list(
            self.name, "skip_columns", skip_columns
        )

    def validate(self, ws, ws_name, **dependencies):
        self.ws_name = ws_name
        self._handle_dependencies(**dependencies, fn=self._check_special_characters)

    def _check_special_characters(self, df):
        # Check the url columns
        if len(self.url_columns) > 0:
            try:
                url_df = df[self.url_columns]
            except KeyError:
                raise KeyError(f"Column(s) {self.url_columns} not found in worksheet {self.ws_name}")
            self.bad_chars_regex = self.url_pattern
            Cell.validate(self, "Invalid URL cells found", "Valid URL cells found", url_df)
        
        # Check the free text columns
        if len(self.free_text_columns) > 0:
            free_text_df = df[self.free_text_columns]
            self.bad_chars_regex = self.free_text_pattern
            Cell.validate(self, "Invalid free text cells found", "Valid free text cells found", free_text_df)

        # Check the rest of the cells
        default_df = df.drop(
            columns=self.url_columns + self.free_text_columns + self.skip_columns,
            errors="ignore",
        )
        self.bad_chars_regex = self.default_pattern
        Cell.validate(self, "General special characters found", "No general special characters found", default_df)


    def not_valid(self, cell):
        """Return True if the cell is not valid (has disallowed special chars). Uses coord for free_text_columns."""
        if not isinstance(cell, str):
            return False

        bad_chars = re.findall(self.bad_chars_regex, cell)
        return len(bad_chars) > 0


# Search for leading or trailing white space
class Cell_Untrimmed_White_Space(Cell):
    def __init__(self):
        Cell.__init__(self)
        
    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_untrimmed_white_space)

    def _check_untrimmed_white_space(self, df):
        # Use the default validate method for Cell
        Cell.validate(self,
            "Leading or trailing white space found",
            "No leading or trailing white space found",
            df)
    
    # Required by Cell
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # This regex matches any white space at the beginning or end of the string
        # Bad_space is a list of all the leading or trailing white space found, but not if the cell is exactly a single space
        if cell == " ":
            bad_space = []
        else:
            bad_space = re.findall(r'^\s|\s$', cell)

        # If there is no leading or trailing white space, return False
        if len(bad_space) == 0:
            return False
        else:
            # The cell had leading or trailing white space
            return True

class Cell_Newlines_Tabs(Cell):
    def __init__(self):
        Cell.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_newlines_tabs)

    def _check_newlines_tabs(self, df):
        # Use the default validate method for Cell
        Cell.validate(self,
            "Newlines, tabs, or vertical tabs found",
            "No newlines, tabs, or vertical tabs found",
            df)
        
    # Required by Cell
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



        
    # Required by Cell
    def not_valid(self, cell):

        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # This regex matches single and double quotes around comma or a tab
        # bad_space is a list of all the matches
        delimiters_in_cells = re.findall(r'[\'\"][,\t][\'\"]', cell)

        # If there are no newlines... etc
        if len(delimiters_in_cells)==0:
            return False
        else:
            # The cell had newlines
            return True

class Cell_Missing_Value_Text(Cell):
    def __init__(self):
        Cell.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_missing_value_text)

    def _check_missing_value_text(self, df):
        # Use the default validate method for Cell
        Cell.validate(self,
            "Text denoting missing values found",
            "No text denoting missing values found",
            df)
        
    # Required by Cell
    def not_valid(self, cell):

        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # This regex matches newlines, tabs, and vertical tabs anywhere
        # bad is a list of all the matches
        bad = re.findall(r'^(?:(no data)|(nd)|(missing)|(missing data)|(na)|(null)|(\-)|(\.)|(\s+)|(_+))$', cell, flags = re.IGNORECASE)

        # If there are matches
        if len(bad)==0:
            return False
        else:
            # The cell had missing data text
            return True

# Search for cells containing only question marks
class Cell_Question_Mark_Only(Cell):
    def __init__(self):
        Cell.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_question_mark_only)

    def _check_question_mark_only(self, df):
        Cell.validate(self, "Cells with just a question mark found", "No question mark cells found", df)
    
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        
        # If the cell is exactly a question mark.
        return cell == "?"
            
class Cell_White_Space_Only(Cell):
    def __init__(self):
        Cell.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_white_space_only)

    def _check_white_space_only(self, df):
        Cell.validate(self, "Cells with white space only found", "No cells with white space only found", df)
    
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

class Cell_Number_Space(Cell):
    def __init__(self):
        Cell.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_number_space)

    def _check_number_space(self, df):
        Cell.validate(self, "Cells with only spaces and numbers found", "No cells with only spaces and numbers found", df)
    
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



class Cell_Dates(Cell):
    def __init__(self, *, date_columns=None, auto_detect_columns=False,
                 format_code="%Y/%m/%d", date_column_threshold=0.8):
        Cell.__init__(self)
        self.date_columns = _validate_optional_str_list(self.name, "date_columns", date_columns)
        _validate_kwarg_type(self.name, "auto_detect_columns", auto_detect_columns, bool)
        self.auto_detect_columns = auto_detect_columns
        _validate_kwarg_type(self.name, "format_code", format_code, str)
        self.format_code = format_code
        _validate_kwarg_type(self.name, "date_column_threshold", date_column_threshold, (int, float))
        if date_column_threshold < 0 or date_column_threshold > 1:
            raise ValueError(
                f"{self.name}: `date_column_threshold` must be between 0 and 1."
            )
        self.threshold = date_column_threshold
        
        
    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_dates)

    def _check_dates(self, df):
        date_df = df[self.date_columns]
        
        if self.auto_detect_columns:
            remaining_df = df.drop(self.date_columns)

            # Coerce all cells to datetimes, 'coerce' will set invalid dates to na
            # 'mixed' will try to parse the dates in any format
            # Note this is usually overly ambitious
            parsed_dates = remaining_df.apply(
                pd.to_datetime, errors="coerce", format="mixed"
            )

            # Check the ratio of valid dates to total dates
            # This will be a series with the column names as the index
            # and the ratio of valid dates as the values
            ratio_dates = parsed_dates.apply(lambda s: s.notna().mean())
            # Get only with more than the threshold of valid dates
            date_cols = ratio_dates[ratio_dates > self.threshold].index
            # Add the new columns to the date dataframe
            date_df = pd.concat([date_df, remaining_df[date_cols]])

        # Validate the date dataframe
        Cell.validate(
            self,
            "Cells with non-ISO dates found",
            "Either no dates or all dates are ISO",
            date_df,
        )


    def not_valid(self, cell):
        if not isinstance(cell, str):
            # This test does not apply
            return False

        # Check if the dates are in the required format
        try:
            # Attempt to parse the date
            cell_date = dt.strptime(cell, self.format_code)
            # If the parsing is successful, the date is valid
            return False
        except ValueError as e:
            # If the parsing is not successful, the date is invalid
            return True


class Cell_Scientific_Notation(Cell):
    # This test fails for "small" exponents like 1e-3
    # Because pandas reads these as floats automatically
    def __init__(self):
        Cell.__init__(self)

    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_scientific_notation)

    def _check_scientific_notation(self, df):
        Cell.validate(self, "Cells with scientific notation found", "No cells with scientific notation found", df)
    
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




_DEFAULT_UNIT_ABBREVS = [
    "nm", "µm", "um", "mm", "cm", "m", "km",
    "in", "ft", "yd", "mi",
    "gal", "L", "mL",
    "g", "kg", "lb", "oz", "mg", "µg", "ug", "ng", "pg",
    "µmol", "umol", "mmol", "mol",
    "s", "sec", "min", "hr", "h", "d", "day", "week", "w", "y", "year", "mo",
]


class Cell_Units(Cell):
    def __init__(self, *, unit_abbreviations=None):
        Cell.__init__(self)
        if unit_abbreviations is None:
            unit_abbreviations = _DEFAULT_UNIT_ABBREVS
        self.unit_abbreviations = _validate_optional_str_list(
            self.name, "unit_abbreviations", unit_abbreviations
        )
        self.units_regex = r'^[+-]?[\d\.]+(' + '|'.join(re.escape(u) for u in self.unit_abbreviations) + r')$'


    def validate(self, *ignore, **dependencies):
        self._handle_dependencies(**dependencies, fn=self._check_units)

    def _check_units(self, df):
        # Check that the grid edges test has been run
        Cell.validate(self, "Cells with units found", "No cells with units found", df)
    
    def not_valid(self, cell):
        
        # Check if the cell is a string
        if not isinstance(cell, str):
            # This test does not apply
            return False
        

        matches = re.fullmatch(self.units_regex, cell.strip())

        # If the cell has no units, return False
        if matches == None:
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