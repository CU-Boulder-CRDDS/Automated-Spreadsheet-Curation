from detectors import Test_Suite
import os


suite = Test_Suite(wb_path = "data/4723768_train_.csv", config_path = "config.example.json")

suite.run()
suite.save(format = "csv")
suite.save(format = "json")
suite.report()

# Get a list of all the files in the data directory
# files = os.listdir("data")


# # For each file, create a test suite
# for file in files:
#     # Create a test suite from the file
#     if file.startswith("."):
#         continue
#     else:
#         suite = Test_Suite(wb_path = f"data/{file}", config_path = "config.example.json")
#         # Run all the tests and outputs
#         suite.run()
#         suite.save(format = "csv")
#         suite.save(format = "json")
#         suite.report()

