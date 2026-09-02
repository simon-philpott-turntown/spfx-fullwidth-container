/**
 * @file ErrorBoundary.tsx
 * @description React Error Boundary to catch any runtime rendering errors gracefully.
 */

import * as React from 'react';

export interface IErrorBoundaryProps {
  children?: React.ReactNode;
  fallbackTitle?: string;
}

export interface IErrorBoundaryState {
  hasError: boolean;
  error?: Error;
  retryCount: number;
}

export class ErrorBoundary extends React.Component<IErrorBoundaryProps, IErrorBoundaryState> {
  private _retryTimer: number | undefined;

  constructor(props: IErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, retryCount: 0 };
  }

  public static getDerivedStateFromError(error: Error): Partial<IErrorBoundaryState> {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    // Automatically perform a silent self-healing retry if initial mount had a race condition
    if (this.state.retryCount < 2) {
      if (this._retryTimer) {
        window.clearTimeout(this._retryTimer);
      }
      this._retryTimer = window.setTimeout(() => {
        this.setState((prev) => ({
          hasError: false,
          error: undefined,
          retryCount: prev.retryCount + 1
        }));
      }, 50);
    }
  }

  public componentWillUnmount(): void {
    if (this._retryTimer) {
      window.clearTimeout(this._retryTimer);
    }
  }

  public render(): React.ReactNode {
    if (this.state.hasError && this.state.retryCount >= 2) {
      return (
        <div style={{
          padding: '24px',
          border: '1px solid #d13438',
          borderRadius: '8px',
          backgroundColor: '#fde7e9',
          color: '#a80000',
          fontFamily: 'Segoe UI, sans-serif'
        }}>
          <h3 style={{ margin: '0 0 8px 0', fontSize: '18px' }}>
            {this.props.fallbackTitle || 'Full-Width Container'}
          </h3>
          <p style={{ margin: '0 0 12px 0', fontSize: '14px' }}>
            {this.state.error?.message || 'An unexpected rendering error occurred.'}
          </p>
          <button
            style={{
              padding: '6px 14px',
              backgroundColor: '#a80000',
              color: '#ffffff',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer'
            }}
            onClick={() => this.setState({ hasError: false, error: undefined, retryCount: 0 })}
          >
            Retry Rendering
          </button>
        </div>
      );
    }

    if (this.state.hasError) {
      // Return null or placeholder while silent auto-recovery runs
      return null;
    }

    return this.props.children;
  }
}
