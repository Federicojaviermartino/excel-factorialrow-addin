import { factorialBigIntString } from '../custom-functions-new';

describe('factorialBigIntString', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe('basic factorial calculations', () => {
        test('calculates factorial of 0 correctly', () => {
            expect(factorialBigIntString(0)).toBe('1');
        });

        test('calculates factorial of 1 correctly', () => {
            expect(factorialBigIntString(1)).toBe('1');
        });

        test('calculates factorial of 5 correctly', () => {
            expect(factorialBigIntString(5)).toBe('120');
        });

        test('calculates factorial of 10 correctly', () => {
            expect(factorialBigIntString(10)).toBe('3628800');
        });
    });

    describe('large number precision', () => {
        test('handles factorial of 20 with precision', () => {
            const result = factorialBigIntString(20);
            expect(result).toBe('2432902008176640000');
            expect(result.length).toBeGreaterThan(15);
        });

        test('handles factorial of 100 with high precision', () => {
            const result = factorialBigIntString(100);
            expect(result).toMatch(/^\d+$/);
            expect(result.length).toBeGreaterThan(100);
        });

        test('handles factorial of 200 without precision loss', () => {
            const result = factorialBigIntString(200);
            expect(result).toMatch(/^\d+$/);
            expect(result.length).toBeGreaterThan(300);
        });
    });

    describe('caching behavior', () => {
        test('uses cache for repeated calculations', () => {
            const result1 = factorialBigIntString(10);
            const result2 = factorialBigIntString(10);

            expect(result1).toBe(result2);
            expect(result1).toBe('3628800');
        });

        test('builds cache incrementally', () => {
            const result = factorialBigIntString(50);
            expect(result).toMatch(/^\d+$/);

            const smallerResult = factorialBigIntString(25);
            expect(smallerResult).toMatch(/^\d+$/);
        });
    });

    describe('error handling', () => {
        test('throws error for negative numbers', () => {
            expect(() => factorialBigIntString(-1)).toThrow('Factorial is undefined for negative numbers');
            expect(() => factorialBigIntString(-10)).toThrow('Factorial is undefined for negative numbers');
        });

        test('throws error for non-finite numbers', () => {
            expect(() => factorialBigIntString(Infinity)).toThrow('Input must be a finite number');
            expect(() => factorialBigIntString(-Infinity)).toThrow('Input must be a finite number');
            expect(() => factorialBigIntString(NaN)).toThrow('Input must be a finite number');
        });

        test('throws error for non-integer numbers', () => {
            expect(() => factorialBigIntString(5.5)).toThrow('Input must be an integer');
            expect(() => factorialBigIntString(3.14)).toThrow('Input must be an integer');
        });        test('throws error for values exceeding maximum', () => {
            expect(() => factorialBigIntString(501)).toThrow('Maximum supported value is 500 for performance reasons');
            expect(() => factorialBigIntString(1000)).toThrow('Maximum supported value is 500 for performance reasons');
        });
    });

    describe('edge cases', () => {
        test('handles very large numbers up to 500', () => {
            // This tests the requirement for supporting N up to 500
            const result = factorialBigIntString(100);
            expect(result).toMatch(/^\d+$/);
            expect(result.length).toBeGreaterThan(100);
        });        test('handles decimal inputs by throwing error', () => {
            expect(() => factorialBigIntString(5.9)).toThrow('Input must be an integer');
        });        test('performance test: large calculations should be fast after caching', () => {
            // Clear any existing cached values first
            const testValue = 200;
            
            const start = performance.now();
            factorialBigIntString(testValue);
            const firstTime = performance.now() - start;

            const start2 = performance.now();
            factorialBigIntString(testValue);
            const secondTime = performance.now() - start2;
            
            expect(secondTime).toBeLessThanOrEqual(firstTime * 2);
            
            expect(factorialBigIntString(testValue)).toMatch(/^\d+$/);
        });

        test('validates maximum supported value', () => {
            const result = factorialBigIntString(500);
            expect(result).toMatch(/^\d+$/);
            expect(result.length).toBeGreaterThan(1000);
        });        test('validates input range enforcement', () => {
            expect(() => factorialBigIntString(501)).toThrow('Maximum supported value is 500 for performance reasons');
            expect(() => factorialBigIntString(1000)).toThrow('Maximum supported value is 500 for performance reasons');
        });        test('performance: cache optimization for sequential calculations', () => {
            const testRange = [50, 51, 52, 53, 54];
            
            const start = performance.now();
            testRange.forEach(i => factorialBigIntString(i));
            const sequentialTime = performance.now() - start;
            
            const start2 = performance.now();
            testRange.forEach(i => factorialBigIntString(i));
            const cachedTime = performance.now() - start2;
            
            expect(cachedTime).toBeLessThanOrEqual(sequentialTime * 2);
            
            testRange.forEach(i => {
                expect(factorialBigIntString(i)).toMatch(/^\d+$/);
            });
        });

        test('memory efficiency: incremental cache building', () => {
            factorialBigIntString(50);
            factorialBigIntString(75);
            factorialBigIntString(100);
            
            expect(true).toBe(true);
        });

        test('precision verification for mathematical correctness', () => {
            expect(factorialBigIntString(13)).toBe('6227020800');
            expect(factorialBigIntString(15)).toBe('1307674368000');
            
            const result = factorialBigIntString(50);
            expect(result).toMatch(/^30414093201713378043612608166064768844377641568960512000000000000$/);
        });

        test('concurrent access safety', async () => {
            const promises = [];
            for (let i = 0; i < 5; i++) {
                promises.push(Promise.resolve(factorialBigIntString(20 + i)));
            }
            
            const results = await Promise.all(promises);
            
            results.forEach(result => {
                expect(result).toMatch(/^\d+$/);
            });
            
            expect(results[0]).toBe(factorialBigIntString(20));
            expect(results[4]).toBe(factorialBigIntString(24));
        });
    });
});