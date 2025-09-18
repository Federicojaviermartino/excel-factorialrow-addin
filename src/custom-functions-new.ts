const factorialCache: Map<number, string> = new Map([[0, "1"], [1, "1"]]);

/**
 * Computes factorial using optimized BigInt with intelligent memoization.
 * Performance: O(1) for cached values, O(n-cached) for new calculations.
 * Memory: Efficient incremental caching, no redundant calculations.
 * @param n - Non-negative integer up to 500
 * @returns String representation maintaining precision
 */
export function factorialBigIntString(n: number): string {
    if (!Number.isFinite(n)) {
        throw new Error("Input must be a finite number");
    }
    if (n < 0) {
        throw new Error("Factorial is undefined for negative numbers");
    }
    if (Math.floor(n) !== n) {
        throw new Error("Input must be an integer");
    }
    if (n > 500) {
        throw new Error("Maximum supported value is 500 for performance reasons");
    }

    if (factorialCache.has(n)) {
        return factorialCache.get(n)!;
    }

    let start = 0;
    let acc = 1n;
    
    for (const [k, val] of factorialCache) {
        if (k <= n && k > start) {
            start = k;
            acc = BigInt(val);
        }
    }

    for (let i = start + 1; i <= n; i++) {
        acc *= BigInt(i);
        factorialCache.set(i, acc.toString());
    }
    
    return factorialCache.get(n)!;
}

/**
 * Reads persisted orientation from OfficeRuntime.storage with error recovery.
 * @returns Promise that resolves to "row" or "column"
 * @internal
 */
async function getOrientation(): Promise<"row" | "column"> {
    try {
        const value = await OfficeRuntime.storage.getItem("orientation");
        return value === "column" ? "column" : "row";
    } catch (error) {
        console.warn('Storage access failed, using default orientation:', error);
        return "row";
    }
}

/**
 * Returns a spilled range of factorials from 0! to N!
 * 
 * @customfunction
 * @param N Maximum factorial value to calculate (integer, >= 0, <= 500)
 * @returns Promise<string[][]> Spilled range with factorials from 0 to N as strings
 * 
 * @example
 * // Row format (default)
 * =TESTVELIXO.FACTORIALROW(5) → [[1, 1, 2, 6, 24, 120]]
 * 
 * @example  
 * // Column format (when orientation = "column")
 * =TESTVELIXO.FACTORIALROW(3) → [[1], [1], [2], [6]]
 * 
 * @performance
 * - Cached values: O(1) retrieval
 * - New calculations: O(n) where n = uncached range
 * - Memory: O(n) space for cache storage
 * 
 * @precision
 * - Supports calculations up to N=500 without precision loss
 * - Uses BigInt arithmetic with string conversion
 * - Maintains mathematical accuracy for large factorials
 */
export async function FACTORIALROW(N: number): Promise<string[][]> {
    try {
        const n = Math.floor(N);
        if (n < 0) {
            throw new Error("N must be a non-negative integer.");
        }

        const values: string[] = new Array(n + 1);
        for (let i = 0; i <= n; i++) {
            values[i] = factorialBigIntString(i);
        }

        const orientation = await getOrientation();

        if (orientation === "column") {
            return values.map(v => [v]);
        } else {
            return [values];
        }
    } catch (error) {
        console.error('FACTORIALROW calculation error:', error);
        throw error;
    }
}

/**
 * FACTORIALROW_WORKER - Web Worker version
 * @customfunction
 * @param N Maximum factorial value to calculate using Web Worker
 * @returns Spilled range with factorials from 0 to N as strings
 */
export async function FACTORIALROW_WORKER(N: number): Promise<string[][]> {
    return new Promise((resolve, reject) => {
        const worker = new Worker('http://localhost:3000/worker.bundle.js');
        worker.postMessage({ n: Math.floor(N) });

        worker.onmessage = async (ev) => {
            worker.terminate();
            if (ev.data.error) {
                reject(new Error(ev.data.error));
            } else {
                const orientation = await getOrientation();
                const values = ev.data.result;

                if (orientation === "column") {
                    resolve(values.map((v: string) => [v]));
                } else {
                    resolve([values]);
                }
            }
        };

        worker.onerror = (error) => {
            worker.terminate();
            reject(error);
        };
    });
}

/**
 * TEST FUNCTION - Simple version
 * @customfunction
 * @param N A number
 * @returns Simple test result
 */
export function TESTFUNC(N: number): number[][] {
    return [[1, 2, 3, N]];
}

/**
 * SIMPLE TEST - Basic function
 * @customfunction
 * @param N A number
 * @returns A simple number
 */
export function SIMPLETEST(N: number): number {
    return N * 2;
}

(function () {
    if (process.env.NODE_ENV !== 'test') {
        console.log('🚀 Starting function registration...');
    }

    const global = (typeof self !== 'undefined' ? self : typeof window !== 'undefined' ? window : globalThis) as any;

    if (!global.TESTVELIXO) {
        global.TESTVELIXO = {};
    }

    global.TESTVELIXO.FACTORIALROW = FACTORIALROW;
    global.TESTVELIXO.FACTORIALROW_WORKER = FACTORIALROW_WORKER;
    global.TESTVELIXO.TESTFUNC = TESTFUNC;

    global.FACTORIALROW = FACTORIALROW;
    global.FACTORIALROW_WORKER = FACTORIALROW_WORKER;
    global.TESTFUNC = TESTFUNC;

    if (process.env.NODE_ENV !== 'test') {
        console.log('✅ Functions registered directly to global scope');
    }
    
    if (typeof window !== 'undefined') {
        (window as any).TESTVELIXO = global.TESTVELIXO;
        (window as any).FACTORIALROW = FACTORIALROW;
        (window as any).FACTORIALROW_WORKER = FACTORIALROW_WORKER;
        (window as any).TESTFUNC = TESTFUNC;
        if (process.env.NODE_ENV !== 'test') {
            console.log('✅ Functions registered to window');
        }
    }

    if (typeof self !== 'undefined') {
        (self as any).TESTVELIXO = global.TESTVELIXO;
        (self as any).FACTORIALROW = FACTORIALROW;
        (self as any).FACTORIALROW_WORKER = FACTORIALROW_WORKER;
        (self as any).TESTFUNC = TESTFUNC;
        if (process.env.NODE_ENV !== 'test') {
            console.log('✅ Functions registered to self');
        }
    }
    
    if (typeof CustomFunctions !== 'undefined') {
        if (CustomFunctions.associate) {
            try {
                CustomFunctions.associate('FACTORIALROW', FACTORIALROW);
                CustomFunctions.associate('FACTORIALROW_WORKER', FACTORIALROW_WORKER);
                CustomFunctions.associate('TESTFUNC', TESTFUNC);
                if (process.env.NODE_ENV !== 'test') {
                    console.log('✅ Functions registered via CustomFunctions.associate');
                }
            } catch (e) {
                console.warn('⚠️ CustomFunctions.associate failed:', e);
            }
        }
    } else {
        if (process.env.NODE_ENV !== 'test') {
            console.warn('⚠️ CustomFunctions not available yet');
        }
    }

    setTimeout(() => {
        if (typeof CustomFunctions !== 'undefined' && CustomFunctions.associate) {
            try {
                CustomFunctions.associate('FACTORIALROW', FACTORIALROW);
                CustomFunctions.associate('FACTORIALROW_WORKER', FACTORIALROW_WORKER);
                CustomFunctions.associate('TESTFUNC', TESTFUNC);
                console.log('✅ Functions registered via delayed CustomFunctions.associate');
            } catch (e) {
                console.warn('⚠️ Delayed CustomFunctions.associate failed:', e);
            }
        }

        global.FACTORIALROW = FACTORIALROW;
        global.FACTORIALROW_WORKER = FACTORIALROW_WORKER;
        global.TESTFUNC = TESTFUNC;
        console.log('✅ Delayed registration complete');
    }, 100);

    setTimeout(() => {
        global.FACTORIALROW = FACTORIALROW;
        global.FACTORIALROW_WORKER = FACTORIALROW_WORKER;
        global.TESTFUNC = TESTFUNC;
        if (typeof CustomFunctions !== 'undefined' && CustomFunctions.associate) {
            try {
                CustomFunctions.associate('FACTORIALROW', FACTORIALROW);
                CustomFunctions.associate('FACTORIALROW_WORKER', FACTORIALROW_WORKER);
                CustomFunctions.associate('TESTFUNC', TESTFUNC);
                console.log('✅ Very delayed CustomFunctions.associate succeeded');
            } catch (e) {
                console.warn('⚠️ Very delayed CustomFunctions.associate failed:', e);
            }
        }
    }, 500);    console.log('🔧 Function registration complete');
    if (process.env.NODE_ENV !== 'test') {
        console.log('📋 TESTVELIXO namespace functions:', Object.keys(global.TESTVELIXO || {}));
        console.log('🔍 Global FACTORIALROW type:', typeof global.FACTORIALROW);
        console.log('🔍 Global FACTORIALROW_WORKER type:', typeof global.FACTORIALROW_WORKER);
        console.log('🔍 Global TESTFUNC type:', typeof global.TESTFUNC);
    }

    if (typeof Office !== 'undefined') {
        Office.onReady(() => {
            console.log('📚 Office.js is ready, functions should be available');
            global.FACTORIALROW = FACTORIALROW;
            global.FACTORIALROW_WORKER = FACTORIALROW_WORKER;
            global.TESTFUNC = TESTFUNC;
            if (typeof CustomFunctions !== 'undefined' && CustomFunctions.associate) {
                try {
                    CustomFunctions.associate('FACTORIALROW', FACTORIALROW);
                    CustomFunctions.associate('FACTORIALROW_WORKER', FACTORIALROW_WORKER);
                    CustomFunctions.associate('TESTFUNC', TESTFUNC);
                    console.log('✅ Office.onReady CustomFunctions.associate succeeded');
                } catch (e) {
                    console.warn('⚠️ Office.onReady CustomFunctions.associate failed:', e);
                }
            }
        });
    }
})();

(function () {
    const global = (typeof self !== 'undefined' ? self : typeof window !== 'undefined' ? window : globalThis) as any;
    global.SIMPLETEST = SIMPLETEST;
    if (!global.TESTVELIXO) global.TESTVELIXO = {};
    global.TESTVELIXO.SIMPLETEST = SIMPLETEST;

    console.log('✅ SIMPLETEST registered:', typeof global.SIMPLETEST);
})();