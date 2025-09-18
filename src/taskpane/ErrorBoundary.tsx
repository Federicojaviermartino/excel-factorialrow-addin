import React, { ErrorInfo, ReactNode } from 'react';

interface ErrorBoundaryState {
    hasError: boolean;
    error?: Error;
}

interface ErrorBoundaryProps {
    children: ReactNode;
}

class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
    constructor(props: ErrorBoundaryProps) {
        super(props);
        this.state = { hasError: false };
    }

    static getDerivedStateFromError(error: Error): ErrorBoundaryState {
        return { hasError: true, error };
    }

    componentDidCatch(error: Error, errorInfo: ErrorInfo): void {
        console.error('Excel Add-in Error:', error, errorInfo);
    }

    render(): ReactNode {
        if (this.state.hasError) {
            return (
                <div style={{
                    padding: 20,
                    backgroundColor: '#fee',
                    border: '2px solid #fcc',
                    borderRadius: 8,
                    color: '#c44',
                    fontFamily: 'system-ui, Arial, sans-serif'
                }}>
                    <h2>⚠️ Something went wrong</h2>
                    <p>The Excel Add-in encountered an unexpected error.</p>
                    <details style={{ marginTop: 16 }}>
                        <summary>Technical Details</summary>
                        <pre style={{
                            fontSize: 12,
                            backgroundColor: '#f8f8f8',
                            padding: 8,
                            borderRadius: 4,
                            overflow: 'auto'
                        }}>
                            {this.state.error?.toString()}
                        </pre>
                    </details>
                    <button
                        onClick={() => window.location.reload()}
                        style={{
                            marginTop: 16,
                            padding: '8px 16px',
                            backgroundColor: '#007fff',
                            color: 'white',
                            border: 'none',
                            borderRadius: 4,
                            cursor: 'pointer'
                        }}
                    >
                        Reload Add-in
                    </button>
                </div>
            );
        }

        return this.props.children;
    }
}

export default ErrorBoundary;