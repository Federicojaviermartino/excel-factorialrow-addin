import { createRoot } from "react-dom/client";
import { App } from "./App";
import ErrorBoundary from "./ErrorBoundary";

const container = document.getElementById("root")!;
const root = createRoot(container);
root.render(
    <ErrorBoundary>
        <App />
    </ErrorBoundary>
);
