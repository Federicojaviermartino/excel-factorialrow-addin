import { render, screen } from '@testing-library/react';
import ErrorBoundary from '../taskpane/ErrorBoundary';

const ThrowError = ({ shouldThrow }: { shouldThrow: boolean }) => {
    if (shouldThrow) {
        throw new Error('Test error');
    }
    return <div>No error</div>;
};

describe('ErrorBoundary', () => {
    const originalError = console.error;
    beforeAll(() => {
        console.error = jest.fn();
    });

    afterAll(() => {
        console.error = originalError;
    });

    test('renders children when there is no error', () => {
        render(
            <ErrorBoundary>
                <ThrowError shouldThrow={false} />
            </ErrorBoundary>
        );

        expect(screen.getByText('No error')).toBeInTheDocument();
    });

    test('renders error UI when there is an error', () => {
        render(
            <ErrorBoundary>
                <ThrowError shouldThrow={true} />
            </ErrorBoundary>
        );

        expect(screen.getByText('⚠️ Something went wrong')).toBeInTheDocument();
        expect(screen.getByText('The Excel Add-in encountered an unexpected error.')).toBeInTheDocument();
        expect(screen.getByText('Reload Add-in')).toBeInTheDocument();
    });

    test('shows technical details when expanded', () => {
        render(
            <ErrorBoundary>
                <ThrowError shouldThrow={true} />
            </ErrorBoundary>
        );

        const detailsElement = screen.getByText('Technical Details');
        expect(detailsElement).toBeInTheDocument();
    });
});