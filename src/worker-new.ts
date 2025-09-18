interface WorkerRequestMessage {
    n: number;
}

interface WorkerResponseMessage {
    result?: string[];
    error?: string;
}

const factorialCache = new Map<number, bigint>([[0, 1n], [1, 1n]]);

function factorial(n: number): bigint {
    if (factorialCache.has(n)) {
        return factorialCache.get(n)!;
    }

    let start = 1;
    let acc = 1n;

    for (const [k, val] of factorialCache) {
        if (k > start && k < n) {
            start = k;
            acc = val;
        }
    }

    for (let i = start + 1; i <= n; i++) {
        acc *= BigInt(i);
        factorialCache.set(i, acc);
    }

    return factorialCache.get(n)!;
}

self.onmessage = (ev: MessageEvent<WorkerRequestMessage>) => {
    try {
        const { n } = ev.data;

        if (n < 0 || !Number.isFinite(n) || Math.floor(n) !== n) {
            self.postMessage({ error: 'N must be a non-negative integer.' });
            return;
        }

        const result: string[] = [];
        for (let i = 0; i <= n; i++) {
            result.push(factorial(i).toString());
        }

        self.postMessage({ result });
    } catch (error) {
        self.postMessage({ error: error instanceof Error ? error.message : 'Unknown error' });
    }
};