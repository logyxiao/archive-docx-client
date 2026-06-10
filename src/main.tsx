import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";

interface ErrorBoundaryState {
  message: string;
}

class ErrorBoundary extends React.Component<React.PropsWithChildren, ErrorBoundaryState> {
  state: ErrorBoundaryState = { message: "" };

  static getDerivedStateFromError(error: unknown): ErrorBoundaryState {
    return { message: error instanceof Error ? error.message : String(error) };
  }

  render() {
    if (this.state.message) {
      return (
        <main className="app-error-shell">
          <section className="app-error-card">
            <h1>程序加载失败</h1>
            <p>{this.state.message}</p>
            <button type="button" className="primary-button" onClick={() => window.location.reload()}>
              重新加载
            </button>
          </section>
        </main>
      );
    }

    return this.props.children;
  }
}

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </React.StrictMode>,
);
