from detectors import Test_Suite
# Create a test suite from the demo notebook
suite = Test_Suite("data/2002-ssaml.xls")

# Run all the tests and outputs
suite.run()
suite.report()