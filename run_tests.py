import os
import sys
import unittest
import time

def run_test_suite():
    print("================================================================================")
    print("  STORE WISE SALES COMPARISON DASHBOARD — TEST SUITE EXECUTION")
    print("  Project ID: P052 | Author: Anu Tyagi | Roll: 2500320140025 | Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Mentor: Mr. Dinesh")
    print("================================================================================")
    
    loader = unittest.TestLoader()
    suite = loader.discover(start_dir="tests", pattern="test_*.py")
    
    start_time = time.perf_counter()
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    duration = time.perf_counter() - start_time
    
    print("\n--------------------------------------------------------------------------------")
    print(f"Total Tests Run: {result.testsRun}")
    print(f"Successful:     {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"Failures:       {len(result.failures)}")
    print(f"Errors:         {len(result.errors)}")
    print(f"Execution Time: {duration:.3f} seconds")
    print("--------------------------------------------------------------------------------")
    
    if result.wasSuccessful():
        print(">>> ALL TECHNICAL TESTS PASSED SUCCESSFULLY! (100% PASS RATE) <<<")
        return 0
    else:
        print(">>> SOME TESTS FAILED! CHECK OUTPUT ABOVE. <<<")
        return 1

if __name__ == "__main__":
    sys.exit(run_test_suite())
