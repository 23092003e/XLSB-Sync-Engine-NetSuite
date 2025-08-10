#!/usr/bin/env python
"""
Auto-Fill Test Suite Runner
Comprehensive test runner for all auto-fill logic tests

This script provides convenient ways to run the complete auto-fill test suite
or specific test categories with proper reporting and coverage analysis.

Usage:
    python run_auto_fill_tests.py                    # Run all tests
    python run_auto_fill_tests.py --core             # Run core logic tests only
    python run_auto_fill_tests.py --integration      # Run integration tests only  
    python run_auto_fill_tests.py --performance      # Run performance tests only
    python run_auto_fill_tests.py --edge-cases       # Run edge case tests only
    python run_auto_fill_tests.py --monitoring       # Run monitoring tests only
    python run_auto_fill_tests.py --coverage         # Run with coverage report
    python run_auto_fill_tests.py --verbose          # Run with verbose output
"""

import sys
import unittest
import argparse
from pathlib import Path
import time
from typing import List, Dict, Any

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

# Import all test modules
import test_auto_fill_logic
import test_auto_fill_integration
import test_auto_fill_performance
import test_auto_fill_edge_cases
import test_auto_fill_monitoring


class AutoFillTestRunner:
    """Test runner for auto-fill functionality tests"""
    
    def __init__(self):
        self.test_modules = {
            'core': test_auto_fill_logic,
            'integration': test_auto_fill_integration,
            'performance': test_auto_fill_performance,
            'edge_cases': test_auto_fill_edge_cases,
            'monitoring': test_auto_fill_monitoring
        }
        
        self.test_descriptions = {
            'core': 'Core auto-fill logic and functionality',
            'integration': 'Integration with processing pipeline and Excel COM',
            'performance': 'Performance benchmarks and scaling tests',
            'edge_cases': 'Edge cases, boundary conditions, and error handling',
            'monitoring': 'Memory optimization and performance monitoring integration'
        }
    
    def run_all_tests(self, verbosity: int = 1, show_coverage: bool = False) -> Dict[str, Any]:
        """Run all auto-fill tests"""
        print("="*70)
        print("XLSB Auto-Fill Logic Test Suite")
        print("="*70)
        print()
        
        results = {}
        total_start_time = time.time()
        
        for category, module in self.test_modules.items():
            print(f"Running {category} tests: {self.test_descriptions[category]}")
            print("-" * 50)
            
            result = self._run_test_module(module, verbosity)
            results[category] = result
            
            print(f"Result: {result['tests_run']} tests, "
                  f"{result['failures']} failures, "
                  f"{result['errors']} errors "
                  f"({result['execution_time']:.2f}s)")
            print()
        
        total_time = time.time() - total_start_time
        
        # Print summary
        self._print_summary(results, total_time)
        
        # Show coverage if requested
        if show_coverage:
            self._show_coverage_summary()
        
        return results
    
    def run_specific_tests(self, categories: List[str], verbosity: int = 1) -> Dict[str, Any]:
        """Run specific test categories"""
        print("="*70)
        print("XLSB Auto-Fill Logic Test Suite - Specific Categories")
        print("="*70)
        print()
        
        results = {}
        total_start_time = time.time()
        
        for category in categories:
            if category not in self.test_modules:
                print(f"Warning: Unknown test category '{category}', skipping...")
                continue
            
            module = self.test_modules[category]
            print(f"Running {category} tests: {self.test_descriptions[category]}")
            print("-" * 50)
            
            result = self._run_test_module(module, verbosity)
            results[category] = result
            
            print(f"Result: {result['tests_run']} tests, "
                  f"{result['failures']} failures, "
                  f"{result['errors']} errors "
                  f"({result['execution_time']:.2f}s)")
            print()
        
        total_time = time.time() - total_start_time
        
        # Print summary
        self._print_summary(results, total_time)
        
        return results
    
    def _run_test_module(self, module, verbosity: int) -> Dict[str, Any]:
        """Run tests from a specific module"""
        start_time = time.time()
        
        # Create test suite from module
        loader = unittest.TestLoader()
        suite = loader.loadTestsFromModule(module)
        
        # Run tests
        runner = unittest.TextTestRunner(
            verbosity=verbosity,
            stream=sys.stdout,
            buffer=True
        )
        
        result = runner.run(suite)
        
        execution_time = time.time() - start_time
        
        return {
            'tests_run': result.testsRun,
            'failures': len(result.failures),
            'errors': len(result.errors),
            'execution_time': execution_time,
            'success_rate': (result.testsRun - len(result.failures) - len(result.errors)) / max(1, result.testsRun)
        }
    
    def _print_summary(self, results: Dict[str, Any], total_time: float):
        """Print test execution summary"""
        print("="*70)
        print("TEST EXECUTION SUMMARY")
        print("="*70)
        
        total_tests = sum(r['tests_run'] for r in results.values())
        total_failures = sum(r['failures'] for r in results.values())
        total_errors = sum(r['errors'] for r in results.values())
        total_success_rate = (total_tests - total_failures - total_errors) / max(1, total_tests)
        
        print(f"Total tests run: {total_tests}")
        print(f"Total failures: {total_failures}")
        print(f"Total errors: {total_errors}")
        print(f"Overall success rate: {total_success_rate:.1%}")
        print(f"Total execution time: {total_time:.2f} seconds")
        print()
        
        # Category breakdown
        print("Category Breakdown:")
        print("-" * 50)
        for category, result in results.items():
            status = "PASS" if result['failures'] == 0 and result['errors'] == 0 else "FAIL"
            print(f"{category:12} | {result['tests_run']:3d} tests | "
                  f"{result['success_rate']:6.1%} | {result['execution_time']:6.2f}s | {status}")
        
        print()
        
        # Performance insights
        if 'performance' in results:
            print("Performance Test Insights:")
            print("-" * 50)
            perf_result = results['performance']
            if perf_result['tests_run'] > 0:
                avg_time_per_test = perf_result['execution_time'] / perf_result['tests_run']
                print(f"Performance tests completed in {perf_result['execution_time']:.2f}s")
                print(f"Average time per performance test: {avg_time_per_test:.3f}s")
                if avg_time_per_test > 5.0:
                    print("[WARNING] Some performance tests took longer than expected")
                else:
                    print("[PASS] Performance tests completed within expected time")
            print()
        
        # Overall assessment
        if total_failures == 0 and total_errors == 0:
            print("[SUCCESS] ALL TESTS PASSED! Auto-fill logic is working correctly.")
        else:
            print("[WARNING] Some tests failed. Please review the failures above.")
            
        if total_success_rate >= 0.95:
            print("[EXCELLENT] Excellent test coverage and reliability (≥95% success rate)")
        elif total_success_rate >= 0.90:
            print("[GOOD] Good test reliability (≥90% success rate)")
        else:
            print("[NEEDS IMPROVEMENT] Test reliability needs improvement (<90% success rate)")
    
    def _show_coverage_summary(self):
        """Show code coverage summary (requires coverage.py)"""
        try:
            import coverage
            print("\n" + "="*70)
            print("CODE COVERAGE ANALYSIS")
            print("="*70)
            print("Coverage analysis requires running with coverage.py:")
            print("  coverage run tests/run_auto_fill_tests.py")
            print("  coverage report")
            print("  coverage html  # for detailed HTML report")
        except ImportError:
            print("\n" + "="*70)
            print("CODE COVERAGE NOT AVAILABLE")
            print("="*70)
            print("Install coverage.py to enable code coverage analysis:")
            print("  pip install coverage")
    
    def run_smoke_tests(self) -> bool:
        """Run a quick smoke test to verify basic functionality"""
        print("Running smoke tests for auto-fill functionality...")
        
        # Import basic test class
        from test_utils import AutoFillTestBase, MockExcelSheet
        
        # Create minimal test
        test_case = AutoFillTestBase()
        test_case.setUp()
        
        try:
            # Test basic auto-fill functionality
            headers = ['Item2', 'Note', 'Factory code']
            sheet = MockExcelSheet()
            rows_added = test_case.processor._add_empty_green_rows(sheet, 1, 3, headers)
            
            if rows_added == 3:
                print("[PASS] Smoke test passed: Basic auto-fill functionality works")
                return True
            else:
                print("[FAIL] Smoke test failed: Auto-fill returned unexpected result")
                return False
                
        except Exception as e:
            print(f"[FAIL] Smoke test failed with error: {e}")
            return False


def main():
    """Main entry point for test runner"""
    parser = argparse.ArgumentParser(description='Auto-Fill Logic Test Suite Runner')
    
    # Test selection arguments
    parser.add_argument('--core', action='store_true', help='Run core logic tests')
    parser.add_argument('--integration', action='store_true', help='Run integration tests')
    parser.add_argument('--performance', action='store_true', help='Run performance tests')
    parser.add_argument('--edge-cases', action='store_true', help='Run edge case tests')
    parser.add_argument('--monitoring', action='store_true', help='Run monitoring tests')
    
    # Output options
    parser.add_argument('--verbose', '-v', action='store_true', help='Verbose output')
    parser.add_argument('--coverage', '-c', action='store_true', help='Show coverage information')
    parser.add_argument('--smoke', action='store_true', help='Run smoke tests only')
    
    args = parser.parse_args()
    
    # Create test runner
    runner = AutoFillTestRunner()
    
    # Handle smoke test
    if args.smoke:
        success = runner.run_smoke_tests()
        sys.exit(0 if success else 1)
    
    # Determine verbosity
    verbosity = 2 if args.verbose else 1
    
    # Determine which tests to run
    categories = []
    if args.core:
        categories.append('core')
    if args.integration:
        categories.append('integration')
    if args.performance:
        categories.append('performance')
    if args.edge_cases:
        categories.append('edge_cases')
    if args.monitoring:
        categories.append('monitoring')
    
    # Run tests
    if categories:
        results = runner.run_specific_tests(categories, verbosity)
    else:
        results = runner.run_all_tests(verbosity, args.coverage)
    
    # Exit with appropriate code
    total_failures = sum(r.get('failures', 0) for r in results.values())
    total_errors = sum(r.get('errors', 0) for r in results.values())
    
    sys.exit(0 if (total_failures == 0 and total_errors == 0) else 1)


if __name__ == '__main__':
    main()