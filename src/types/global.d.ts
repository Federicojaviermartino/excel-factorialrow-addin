/// <reference types="jest" />
/// <reference types="@testing-library/jest-dom" />

declare global {
    const Excel: {
        run<T>(callback: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
        CalculationType: {
            full: string;
        };
    };

    const Office: {
        onReady(callback?: (info: { host: string; platform: string }) => void): Promise<{ host: string; platform: string }>;
    };

    const OfficeRuntime: {
        storage: {
            getItem(key: string): Promise<string | null>;
            setItem(key: string, value: string): Promise<void>;
        };
    };

    const CustomFunctions: {
        associate(id: string, functionObject: Function): void;
    };

    namespace Excel {
        interface RequestContext {
            workbook: {
                application: {
                    calculate(calculationType: string): void;
                };
            };
            sync(): Promise<void>;
        }
    }

    namespace jest {
        interface Matchers<R> {
            toBeInTheDocument(): R;
            toBeChecked(): R;
            toBeDisabled(): R;
        }
    }
}

export { };
