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
}

export class ErrorBoundary extends React.Component<IErrorBoundaryProps, IErrorBoundaryState> {
  constructor(props: IErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false };
  }

  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    // Log error internally
  }

  public render(): React.ReactNode {
    if (this.state.hasError) {
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
            {this.props.fallbackTitle || 'Full-Width Container Error'}
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
            onClick={() => this.setState({ hasError: false, error: undefined })}
          >
            Retry Rendering
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
