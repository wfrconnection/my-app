import { useEffect, useState } from 'react'
import { FailedflowsService } from './generated/services/FailedflowsService'
import type { FailedflowsRead } from './generated/models/FailedflowsModel'
import './App.css'

const columns = [
  'ID',
  'Title',
  'Flow name',
  'Flow step',
  'Message',
  'Status',
  'Environment',
  'Created by',
  'Created',
  'Handle status',
]

function getErrorMessage(error: unknown) {
  if (error instanceof Error && error.message) {
    return error.message
  }

  if (typeof error === 'string' && error.trim()) {
    return error
  }

  return 'We couldn’t load the failed flow items. Please try again.'
}

function getStatusTone(status?: string) {
  const normalizedStatus = status?.toLowerCase() ?? ''

  if (/fail|error/.test(normalizedStatus)) {
    return 'status-pill--danger'
  }

  if (/success|complete|resolved|handled/.test(normalizedStatus)) {
    return 'status-pill--success'
  }

  if (/pending|progress|open|new/.test(normalizedStatus)) {
    return 'status-pill--warning'
  }

  return 'status-pill--neutral'
}

function formatCreatedDate(value?: string) {
  if (!value) {
    return '—'
  }

  const date = new Date(value)

  if (Number.isNaN(date.getTime())) {
    return value
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(date)
}

function RefreshIcon() {
  return (
    <svg aria-hidden="true" viewBox="0 0 20 20" fill="none">
      <path
        d="M16.2 8A6.5 6.5 0 0 0 4.7 5.2L3.2 6.7M3.2 6.7V3.5M3.2 6.7h3.2M3.8 12a6.5 6.5 0 0 0 11.5 2.8l1.5-1.5m0 0v3.2m0-3.2h-3.2"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function App() {
  const [records, setRecords] = useState<FailedflowsRead[]>([])
  const [isLoading, setIsLoading] = useState(true)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const [refreshKey, setRefreshKey] = useState(0)

  useEffect(() => {
    let isMounted = true

    async function loadRecords() {
      setIsLoading(true)
      setErrorMessage(null)

      try {
        const result = await FailedflowsService.getAll()

        if (!result.success) {
          throw new Error(
            result.error?.message ??
              'The SharePoint request could not be completed.',
          )
        }

        if (isMounted) {
          setRecords(result.data)
        }
      } catch (error) {
        if (isMounted) {
          setErrorMessage(getErrorMessage(error))
        }
      } finally {
        if (isMounted) {
          setIsLoading(false)
        }
      }
    }

    void loadRecords()

    return () => {
      isMounted = false
    }
  }, [refreshKey])

  return (
    <main className="app-shell">
      <header className="page-header">
        <div>
          <p className="eyebrow">Operations overview</p>
          <h1>Failed flows</h1>
          <p className="page-description">
            Review flow errors and track their handling status.
          </p>
        </div>
        <button
          className="refresh-button"
          type="button"
          onClick={() => setRefreshKey((key) => key + 1)}
          disabled={isLoading}
        >
          <RefreshIcon />
          <span>Refresh</span>
        </button>
      </header>

      <section className="data-card" aria-label="Failed flow items">
        <div className="data-card-header">
          <div>
            <h2>All items</h2>
            <p>
              {isLoading
                ? 'Fetching the latest items from SharePoint…'
                : `${records.length} ${records.length === 1 ? 'item' : 'items'}`}
            </p>
          </div>
          <span className="source-badge">
            <span className="source-indicator" />
            SharePoint list
          </span>
        </div>

        {errorMessage ? (
          <div className="feedback-state" role="alert">
            <div className="feedback-icon feedback-icon--error" aria-hidden="true">
              !
            </div>
            <h2>We couldn’t load the items</h2>
            <p>{errorMessage}</p>
            <button
              className="retry-button"
              type="button"
              onClick={() => setRefreshKey((key) => key + 1)}
            >
              <RefreshIcon />
              Try again
            </button>
          </div>
        ) : isLoading ? (
          <div
            className="table-scroll"
            aria-label="Loading failed flow items"
            aria-busy="true"
          >
            <table className="records-table">
              <thead>
                <tr>
                  {columns.map((column) => (
                    <th key={column} scope="col">
                      {column}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Array.from({ length: 6 }, (_, rowIndex) => (
                  <tr className="skeleton-row" key={rowIndex}>
                    {columns.map((column, columnIndex) => (
                      <td key={column}>
                        <span
                          className={`skeleton-line skeleton-line--${columnIndex % 4}`}
                        />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : records.length === 0 ? (
          <div className="feedback-state">
            <div className="feedback-icon feedback-icon--empty" aria-hidden="true">
              <span />
            </div>
            <h2>No failed flows yet</h2>
            <p>When a flow failure is recorded, it’ll appear here.</p>
          </div>
        ) : (
          <div className="table-scroll">
            <table className="records-table">
              <thead>
                <tr>
                  {columns.map((column) => (
                    <th key={column} scope="col">
                      {column}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {records.map((record, index) => (
                  <tr key={record.ID ?? `record-${index}`}>
                    <td className="id-cell">{record.ID ?? '—'}</td>
                    <td className="title-cell" title={record.Title}>
                      {record.Title || '—'}
                    </td>
                    <td title={record.Flowname}>{record.Flowname || '—'}</td>
                    <td title={record.Step}>{record.Step || '—'}</td>
                    <td className="message-cell" title={record.Message}>
                      {record.Message || '—'}
                    </td>
                    <td>
                      {record.Status ? (
                        <span
                          className={`status-pill ${getStatusTone(record.Status)}`}
                        >
                          <span className="status-dot" />
                          {record.Status}
                        </span>
                      ) : (
                        '—'
                      )}
                    </td>
                    <td title={record.Environment}>{record.Environment || '—'}</td>
                    <td title={record.Author?.DisplayName}>
                      {record.Author?.DisplayName || '—'}
                    </td>
                    <td className="date-cell">
                      {formatCreatedDate(record.Created)}
                    </td>
                    <td>
                      {record.HandleStatus ? (
                        <span
                          className={`status-pill ${getStatusTone(record.HandleStatus)}`}
                        >
                          <span className="status-dot" />
                          {record.HandleStatus}
                        </span>
                      ) : (
                        '—'
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {!isLoading && !errorMessage && records.length > 0 && (
          <div className="table-footer">
            <span>
              Showing <strong>{records.length}</strong>{' '}
              {records.length === 1 ? 'item' : 'items'}
            </span>
            <span className="scroll-hint">Scroll horizontally to see all columns</span>
          </div>
        )}
      </section>
    </main>
  )
}

export default App
