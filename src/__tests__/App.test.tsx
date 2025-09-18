import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { act } from 'react';
import { App } from '../taskpane/App';

const mockGetItem = jest.fn();
const mockSetItem = jest.fn();

const mockExcelRun = jest.fn();
const mockCalculate = jest.fn();

beforeAll(() => {
    (global as any).OfficeRuntime = {
        storage: {
            getItem: mockGetItem,
            setItem: mockSetItem
        }
    };

    (global as any).Excel = {
        run: mockExcelRun,
        CalculationType: {
            full: 'Full'
        }
    };
});

describe('App Component', () => {
    beforeEach(() => {
        jest.clearAllMocks();

        mockGetItem.mockResolvedValue('row');
        mockSetItem.mockResolvedValue(undefined);        mockExcelRun.mockImplementation((callback: any) =>
            callback({
                workbook: {
                    application: {
                        calculate: mockCalculate
                    }
                },
                sync: jest.fn().mockResolvedValue(undefined)
            })
        );
    });    describe('initial render', () => {
        test('renders orientation selector', async () => {
            await act(async () => {
                render(<App />);
            });

            expect(screen.getByText('Factorial: Orientation')).toBeInTheDocument();
            expect(screen.getByText(/Select how the results of/)).toBeInTheDocument();
            expect(screen.getByRole('radiogroup')).toBeInTheDocument();
        });        test('renders row and column radio buttons', async () => {
            await act(async () => {
                render(<App />);
            });

            const rowRadio = screen.getByRole('radio', { name: /row/i });
            const columnRadio = screen.getByRole('radio', { name: /column/i });

            expect(rowRadio).toBeInTheDocument();
            expect(columnRadio).toBeInTheDocument();
        });        test('loads initial orientation from storage', async () => {
            mockGetItem.mockResolvedValue('column');

            await act(async () => {
                render(<App />);
            });

            await waitFor(() => {
                expect(mockGetItem).toHaveBeenCalledWith('orientation');
            });
        });
    });

    describe('orientation selection', () => {
        test('defaults to row orientation', async () => {
            mockGetItem.mockResolvedValue(null);

            render(<App />);

            await waitFor(() => {
                const rowRadio = screen.getByLabelText(/row/i);
                expect(rowRadio).toBeChecked();
            });
        });

        test('selects column when stored preference is column', async () => {
            mockGetItem.mockResolvedValue('column');

            render(<App />);

            await waitFor(() => {
                const columnRadio = screen.getByLabelText(/column/i);
                expect(columnRadio).toBeChecked();
            });
        });

        test('changes orientation when clicking radio button', async () => {
            render(<App />);

            const columnRadio = screen.getByLabelText(/column/i);
            fireEvent.click(columnRadio);

            await waitFor(() => {
                expect(mockSetItem).toHaveBeenCalledWith('orientation', 'column');
                expect(mockExcelRun).toHaveBeenCalled();
                expect(mockCalculate).toHaveBeenCalledWith('Full');
            });
        });

        test('disables buttons while processing', async () => {
            render(<App />);

            const columnRadio = screen.getByLabelText(/column/i);
            const rowRadio = screen.getByLabelText(/row/i);

            fireEvent.click(columnRadio);

            expect(columnRadio).toBeDisabled();
            expect(rowRadio).toBeDisabled();

            await waitFor(() => {
                expect(columnRadio).not.toBeDisabled();
                expect(rowRadio).not.toBeDisabled();
            });
        });
    });    describe('error handling', () => {
        test('handles storage errors gracefully', async () => {
            mockGetItem.mockRejectedValue(new Error('Storage error'));

            await act(async () => {
                render(<App />);
            });

            await waitFor(() => {
                const rowRadio = screen.getByLabelText(/row/i);
                expect(rowRadio).toBeChecked(); // Should default to row
            });
        });        test('handles Excel calculation errors gracefully', async () => {
            mockExcelRun.mockRejectedValue(new Error('Excel error'));

            await act(async () => {
                render(<App />);
            });

            const columnRadio = screen.getByLabelText(/column/i);
            
            await act(async () => {
                fireEvent.click(columnRadio);
            });

            await waitFor(() => {
                expect(mockSetItem).toHaveBeenCalledWith('orientation', 'column');
                expect(columnRadio).not.toBeChecked();
                const rowRadio = screen.getByLabelText(/row/i);
                expect(rowRadio).toBeChecked();
            });
        });
    });    describe('usage instructions', () => {
        test('displays usage instructions', async () => {
            await act(async () => {
                render(<App />);
            });

            expect(screen.getByText('Usage')).toBeInTheDocument();
            expect(screen.getByText(/TESTVELIXO.FACTORIALROW\(10\)/)).toBeInTheDocument();
            expect(screen.getByText(/Use the buttons above to toggle between/)).toBeInTheDocument();
        });

        test('displays storage information', async () => {
            await act(async () => {
                render(<App />);
            });

            expect(screen.getByText(/OfficeRuntime.storage/)).toBeInTheDocument();
        });
    });
});