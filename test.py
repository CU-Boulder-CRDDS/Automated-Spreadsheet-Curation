from detectors import Test_Suite
# Create a test suite from the demo notebook
suite = Test_Suite("data/98117072-28e4-450a-8dba-472b5a8315cc.xlsx")

# Run all the tests and outputs
suite.run()
suite.save(format = "csv")