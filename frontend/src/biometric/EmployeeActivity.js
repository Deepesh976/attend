import React, { useState, useEffect, useMemo } from 'react';
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import { FaChevronDown, FaChevronUp, FaTrashAlt, FaUpload, FaSearch, FaFilter, FaSort, FaCalendarAlt, FaUser, FaClock, FaChartBar, FaTimes } from 'react-icons/fa';

const EmployeeActivity = () => {
  const [allData, setAllData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [selectedFile, setSelectedFile] = useState(null);
  const [expandedRow, setExpandedRow] = useState(null);

  // Enhanced Filter States
  const [filters, setFilters] = useState({
    search: '',
    empId: '',
    empName: '',
    status: '',
    startDate: '',
    endDate: '',
  });

  // Pagination and Display States
  const [currentPage, setCurrentPage] = useState(1);
  const [recordsPerPage, setRecordsPerPage] = useState(25);
  const [sortConfig, setSortConfig] = useState({
    key: 'date',
    direction: 'desc'
  });

  // Display toggles
  const [showAdvancedFilters, setShowAdvancedFilters] = useState(false);
  const [showSummaryPanel, setShowSummaryPanel] = useState(true);
  
  // Monthly Summary Popup
  const [showMonthlySummary, setShowMonthlySummary] = useState(false);
  const [monthlySummaryData, setMonthlySummaryData] = useState([]);
  const [selectedEmployee, setSelectedEmployee] = useState(null);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    setLoading(true);
    try {
      console.log('🔄 Fetching all activity data...');
      const res = await fetch('/api/activity');
      const json = await res.json();
      const activities = json.activities || json || [];
      
      console.log(`📊 Fetched ${activities.length} total records`);
      setAllData(activities);
      setCurrentPage(1); // Reset to first page when data changes
      
      if (activities.length === 0) {
        toast.info('No activity records found. Upload an Excel file to get started.');
      }
    } catch (error) {
      console.error('Fetch error:', error);
      toast.error('Failed to fetch activity data');
      setAllData([]);
    } finally {
      setLoading(false);
    }
  };

  // Enhanced filtering logic
  const filteredData = useMemo(() => {
    let filtered = [...allData];

    // Text search across multiple fields
    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      filtered = filtered.filter(item =>
        item.empId?.toLowerCase().includes(searchLower) ||
        item.empName?.toLowerCase().includes(searchLower) ||
        item.status?.toLowerCase().includes(searchLower) ||
        item.shift?.toLowerCase().includes(searchLower)
      );
    }

    // Employee ID filter
    if (filters.empId) {
      filtered = filtered.filter(item =>
        item.empId?.toLowerCase().includes(filters.empId.toLowerCase())
      );
    }

    // Employee Name filter
    if (filters.empName) {
      filtered = filtered.filter(item =>
        item.empName?.toLowerCase().includes(filters.empName.toLowerCase())
      );
    }

    // Status filter
    if (filters.status) {
      filtered = filtered.filter(item => item.status === filters.status);
    }

    // Date range filter
    if (filters.startDate) {
      filtered = filtered.filter(item => 
        new Date(item.date) >= new Date(filters.startDate)
      );
    }
    if (filters.endDate) {
      filtered = filtered.filter(item => 
        new Date(item.date) <= new Date(filters.endDate)
      );
    }

    return filtered;
  }, [allData, filters]);

  // Enhanced sorting logic
  const sortedData = useMemo(() => {
    if (!sortConfig.key) return filteredData;

    return [...filteredData].sort((a, b) => {
      let aVal = a[sortConfig.key];
      let bVal = b[sortConfig.key];

      // Handle date sorting
      if (sortConfig.key === 'date') {
        aVal = new Date(aVal);
        bVal = new Date(bVal);
      }

      // Handle numeric sorting
      if (typeof aVal === 'string' && aVal.match(/^\d+$/)) {
        aVal = parseInt(aVal);
        bVal = parseInt(bVal);
      }

      // Handle time sorting (HH:MM:SS format)
      if (typeof aVal === 'string' && aVal.match(/^\d{2}:\d{2}(:\d{2})?$/)) {
        const timeToMinutes = (time) => {
          const parts = time.split(':');
          return parseInt(parts[0]) * 60 + parseInt(parts[1]);
        };
        aVal = timeToMinutes(aVal);
        bVal = timeToMinutes(bVal);
      }

      if (aVal < bVal) {
        return sortConfig.direction === 'asc' ? -1 : 1;
      }
      if (aVal > bVal) {
        return sortConfig.direction === 'asc' ? 1 : -1;
      }
      return 0;
    });
  }, [filteredData, sortConfig]);

  // Pagination logic
  const totalPages = Math.ceil(sortedData.length / recordsPerPage);
  const startIndex = (currentPage - 1) * recordsPerPage;
  const endIndex = startIndex + recordsPerPage;
  const currentPageData = sortedData.slice(startIndex, endIndex);

  // Summary statistics
  const summary = useMemo(() => {
    const stats = {
      total: filteredData.length,
      present: filteredData.filter(d => d.status === 'P').length,
      halfDay: filteredData.filter(d => d.status === '½P').length,
      absent: filteredData.filter(d => d.status === 'A').length,
      weeklyOff: filteredData.filter(d => d.status === 'WO').length,
      employees: [...new Set(filteredData.map(d => d.empId))].length,
      dateRange: {
        start: null,
        end: null
      }
    };

    if (filteredData.length > 0) {
      const dates = filteredData.map(d => new Date(d.date)).sort();
      stats.dateRange.start = dates[0];
      stats.dateRange.end = dates[dates.length - 1];
    }

    return stats;
  }, [filteredData]);

  const handleSort = (key) => {
    setSortConfig(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };

  const handleFilterChange = (key, value) => {
    setFilters(prev => ({ ...prev, [key]: value }));
    setCurrentPage(1); // Reset to first page when filter changes
  };

  const clearFilters = () => {
    setFilters({
      search: '',
      empId: '',
      empName: '',
      status: '',
      startDate: '',
      endDate: '',
    });
    setCurrentPage(1);
  };

  // Fetch monthly summary for a specific employee
  const fetchMonthlySummary = async (empId, empName) => {
    setLoading(true);
    try {
      console.log(`📊 Fetching monthly summary for employee: ${empId}`);
      const res = await fetch(`/api/monthly-summary/employee/${empId}`);
      const json = await res.json();
      
      if (res.ok && json.success) {
        console.log(`✅ Retrieved ${json.count} monthly summaries for ${empId}`);
        setMonthlySummaryData(json.data);
        setSelectedEmployee({ empId, empName });
        setShowMonthlySummary(true);
        
        if (json.count === 0) {
          toast.info(`No monthly summary found for employee ${empId}`);
        }
      } else {
        toast.error(json.message || 'Failed to fetch monthly summary');
        setMonthlySummaryData([]);
      }
    } catch (error) {
      console.error('Monthly summary fetch error:', error);
      toast.error('Failed to fetch monthly summary');
      setMonthlySummaryData([]);
    } finally {
      setLoading(false);
    }
  };

  const closeMonthlySummaryPopup = () => {
    setShowMonthlySummary(false);
    setMonthlySummaryData([]);
    setSelectedEmployee(null);
  };

  const handleExcelUpload = async () => {
    if (!selectedFile) {
      toast.warn("Please select a file.");
      return;
    }

    const formData = new FormData();
    formData.append('file', selectedFile);

    setLoading(true);
    try {
      const res = await fetch('/api/activity/upload-excel', {
        method: 'POST',
        body: formData,
      });

      const result = await res.json();

      if (res.ok) {
        toast.success(result.message || 'Data uploaded successfully');
        if (result.skippedRows && result.skippedRows.length > 0) {
          toast.info(`${result.skippedRows.length} rows were skipped`);
        }
        setSelectedFile(null);
        await fetchData();
        clearFilters(); // Clear filters after upload
      } else {
        toast.error(result.message || 'Upload failed');
      }
    } catch (error) {
      console.error('Upload error:', error);
      toast.error('Server error during upload');
    } finally {
      setLoading(false);
    }
  };

  const deleteAllData = async () => {
    if (!window.confirm('Are you sure you want to delete all activity data?')) return;

    setLoading(true);
    try {
      const res = await fetch('/api/activity', { method: 'DELETE' });
      const result = await res.json();
      
      if (res.ok) {
        toast.success(result.message || 'All data deleted');
        setAllData([]);
      } else {
        toast.error(result.message || 'Failed to delete data');
      }
    } catch (error) {
      console.error('Delete error:', error);
      toast.error('Server error during deletion');
    } finally {
      setLoading(false);
    }
  };

  const statusColor = (status) => ({
    'P': '#28a745',
    '½P': '#ffc107',
    'A': '#dc3545',
    'WO': '#6c757d',
  }[status] || '#333');

  const SortableHeader = ({ column, children }) => (
    <th 
      style={{...styles.th, cursor: 'pointer'}} 
      onClick={() => handleSort(column)}
      title={`Sort by ${children}`}
    >
      <div style={styles.headerContent}>
        {children}
        {sortConfig.key === column && (
          <FaSort style={{
            marginLeft: 4,
            transform: sortConfig.direction === 'desc' ? 'rotate(180deg)' : 'none'
          }} />
        )}
      </div>
    </th>
  );

  return (
    <div style={styles.wrapper}>
      <ToastContainer position="top-right" />
      
      {loading && (
        <div style={styles.loadingOverlay}>
          <div style={styles.loadingSpinner}>
            🔄 Loading...
          </div>
        </div>
      )}

      {/* Header Section */}
      <div style={styles.header}>
        <h2 style={styles.title}>
          <FaClock style={styles.titleIcon} />
          Employee Activity Management
        </h2>
        
        {/* Quick Stats */}
        {showSummaryPanel && (
          <div style={styles.quickStats}>
            <div style={styles.statCard}>
              <div style={styles.statNumber}>{summary.total}</div>
              <div style={styles.statLabel}>Total Records</div>
            </div>
            <div style={styles.statCard}>
              <div style={styles.statNumber}>{summary.employees}</div>
              <div style={styles.statLabel}>Employees</div>
            </div>
            <div style={styles.statCard}>
              <div style={styles.statNumber}>{summary.present}</div>
              <div style={styles.statLabel}>Present</div>
            </div>
            <div style={styles.statCard}>
              <div style={styles.statNumber}>{summary.absent}</div>
              <div style={styles.statLabel}>Absent</div>
            </div>
          </div>
        )}
      </div>

      {/* Controls Section */}
      <div style={styles.controlsSection}>
        {/* Primary Controls */}
        <div style={styles.primaryControls}>
          <div style={styles.searchContainer}>
            <FaSearch style={styles.searchIcon} />
            <input
              type="text"
              placeholder="Search by Employee ID, Name, Status..."
              value={filters.search}
              onChange={(e) => handleFilterChange('search', e.target.value)}
              style={styles.searchInput}
            />
          </div>
          
          <button 
            onClick={() => setShowAdvancedFilters(!showAdvancedFilters)}
            style={styles.filterToggle}
          >
            <FaFilter /> Advanced Filters
          </button>

          <label style={styles.uploadButton}>
            <FaUpload /> Upload Excel
            <input
              type="file"
              hidden
              accept=".xlsx,.xls"
              onChange={(e) => setSelectedFile(e.target.files[0])}
            />
          </label>

          <button onClick={deleteAllData} style={styles.deleteButton}>
            <FaTrashAlt /> Delete All
          </button>
        </div>

        {/* Advanced Filters */}
        {showAdvancedFilters && (
          <div style={styles.advancedFilters}>
            <div style={styles.filterRow}>
              <div style={styles.filterGroup}>
                <label>Employee ID:</label>
                <input
                  type="text"
                  value={filters.empId}
                  onChange={(e) => handleFilterChange('empId', e.target.value)}
                  style={styles.filterInput}
                  placeholder="Filter by Employee ID"
                />
              </div>

              <div style={styles.filterGroup}>
                <label>Employee Name:</label>
                <input
                  type="text"
                  value={filters.empName}
                  onChange={(e) => handleFilterChange('empName', e.target.value)}
                  style={styles.filterInput}
                  placeholder="Filter by Employee Name"
                />
              </div>

              <div style={styles.filterGroup}>
                <label>Status:</label>
                <select
                  value={filters.status}
                  onChange={(e) => handleFilterChange('status', e.target.value)}
                  style={styles.filterSelect}
                >
                  <option value="">All Status</option>
                  <option value="P">Present</option>
                  <option value="½P">Half Day</option>
                  <option value="A">Absent</option>
                  <option value="WO">Weekly Off</option>
                </select>
              </div>
            </div>

            <div style={styles.filterRow}>
              <div style={styles.filterGroup}>
                <label>Start Date:</label>
                <input
                  type="date"
                  value={filters.startDate}
                  onChange={(e) => handleFilterChange('startDate', e.target.value)}
                  style={styles.filterInput}
                />
              </div>

              <div style={styles.filterGroup}>
                <label>End Date:</label>
                <input
                  type="date"
                  value={filters.endDate}
                  onChange={(e) => handleFilterChange('endDate', e.target.value)}
                  style={styles.filterInput}
                />
              </div>

              <div style={styles.filterGroup}>
                <button onClick={clearFilters} style={styles.clearButton}>
                  Clear All Filters
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Selected File Display */}
        {selectedFile && (
          <div style={styles.uploadInfo}>
            <span>📎 <strong>{selectedFile.name}</strong></span>
            <button onClick={handleExcelUpload} style={styles.confirmButton}>
              Confirm Upload
            </button>
            <button onClick={() => setSelectedFile(null)} style={styles.cancelButton}>
              Cancel
            </button>
          </div>
        )}
      </div>

      {/* Results Summary */}
      <div style={styles.resultsInfo}>
        <div style={styles.resultsText}>
          Showing {currentPageData.length} of {sortedData.length} records
          {sortedData.length !== allData.length && ` (filtered from ${allData.length} total)`}
        </div>
        
        <div style={styles.paginationControls}>
          <span>Records per page:</span>
          <select
            value={recordsPerPage}
            onChange={(e) => {
              setRecordsPerPage(parseInt(e.target.value));
              setCurrentPage(1);
            }}
            style={styles.pageSelect}
          >
            <option value={10}>10</option>
            <option value={25}>25</option>
            <option value={50}>50</option>
            <option value={100}>100</option>
            <option value={sortedData.length}>All</option>
          </select>
        </div>
      </div>

      {/* Data Table */}
      <div style={styles.tableContainer}>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>S.No</th>
              <SortableHeader column="date">Date</SortableHeader>
              <SortableHeader column="empId">Emp ID</SortableHeader>
              <SortableHeader column="empName">Emp Name</SortableHeader>
              <SortableHeader column="shift">Shift</SortableHeader>
              <SortableHeader column="timeInActual">In Time</SortableHeader>
              <SortableHeader column="timeOutActual">Out Time</SortableHeader>
              <SortableHeader column="lateBy">Late By</SortableHeader>
              <SortableHeader column="earlyBy">Early By</SortableHeader>
              <SortableHeader column="ot">OT</SortableHeader>
              <SortableHeader column="duration">Duration</SortableHeader>
              <SortableHeader column="totalDuration">T Duration</SortableHeader>
              <SortableHeader column="status">Status</SortableHeader>
              <th style={styles.th}>Monthly Summary</th>
            </tr>
          </thead>
          <tbody>
            {currentPageData.map((item, idx) => (
              <React.Fragment key={item._id || idx}>
                <tr 
                  style={{
                    ...styles.row,
                    backgroundColor: idx % 2 === 0 ? '#f8f9fa' : '#ffffff'
                  }}
                >
                  <td style={styles.td}>{startIndex + idx + 1}</td>
                  <td style={styles.td}>{new Date(item.date).toLocaleDateString()}</td>
                  <td style={styles.td}>{item.empId}</td>
                  <td style={styles.td}>{item.empName}</td>
                  <td style={styles.td}>{item.shift}</td>
                  <td style={styles.td}>{item.timeInActual}</td>
                  <td style={styles.td}>{item.timeOutActual}</td>
                  <td style={styles.td}>{item.lateBy}</td>
                  <td style={styles.td}>{item.earlyBy}</td>
                  <td style={styles.td}>{item.ot}</td>
                  <td style={styles.td}>{item.duration}</td>
                  <td style={styles.td}>{item.totalDuration}</td>
                  <td style={styles.td}>
                    <span style={{
                      backgroundColor: statusColor(item.status),
                      color: '#fff',
                      padding: '4px 8px',
                      borderRadius: 4,
                      fontSize: '12px',
                      fontWeight: 'bold'
                    }}>
                      {item.status}
                    </span>
                  </td>
                  <td style={styles.td}>
                    <button
                      onClick={() => fetchMonthlySummary(item.empId, item.empName)}
                      style={styles.summaryButton}
                      title={`View monthly summary for ${item.empName}`}
                    >
                      <FaChartBar />
                    </button>
                  </td>
                </tr>
              </React.Fragment>
            ))}
          </tbody>
        </table>

        {currentPageData.length === 0 && (
          <div style={styles.noData}>
            <FaUser style={styles.noDataIcon} />
            <h3>No Records Found</h3>
            <p>
              {allData.length === 0 
                ? "Upload an Excel file to get started" 
                : "Try adjusting your filters to see more results"
              }
            </p>
          </div>
        )}
      </div>

      {/* Pagination */}
      {totalPages > 1 && (
        <div style={styles.pagination}>
          <button
            onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
            disabled={currentPage === 1}
            style={styles.pageButton}
          >
            Previous
          </button>
          
          <span style={styles.pageInfo}>
            Page {currentPage} of {totalPages}
          </span>
          
          <button
            onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
            disabled={currentPage === totalPages}
            style={styles.pageButton}
          >
            Next
          </button>
        </div>
      )}

      {/* Monthly Summary Popup */}
      {showMonthlySummary && (
        <div style={styles.popupOverlay}>
          <div style={styles.popupContent}>
            <div style={styles.popupHeader}>
              <h3 style={styles.popupTitle}>
                <FaChartBar style={styles.popupIcon} />
                Monthly Summary - {selectedEmployee?.empName} ({selectedEmployee?.empId})
              </h3>
              <button onClick={closeMonthlySummaryPopup} style={styles.closeButton}>
                <FaTimes />
              </button>
            </div>
            
            <div style={styles.popupBody}>
              {monthlySummaryData.length === 0 ? (
                <div style={styles.noSummaryData}>
                  <FaChartBar style={styles.noDataIcon} />
                  <h4>No Monthly Summary Found</h4>
                  <p>No monthly summary data available for this employee.</p>
                </div>
              ) : (
                <div style={styles.summaryList}>
                  {monthlySummaryData.map((summary, index) => (
                    <div key={index} style={styles.summaryCard}>
                      <div style={styles.summaryHeader}>
                        <h4>{summary.monthName} {summary.year}</h4>
                        <span style={styles.summaryBadge}>
                          {summary.totalPresent + summary.totalAbsent} days
                        </span>
                      </div>
                      
                      <div style={styles.summaryGrid}>
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Present Days</span>
                          <span style={{...styles.summaryValue, color: '#28a745'}}>
                            {summary.totalPresent}
                          </span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Absent Days</span>
                          <span style={{...styles.summaryValue, color: '#dc3545'}}>
                            {summary.totalAbsent}
                          </span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Leave Taken</span>
                          <span style={styles.summaryValue}>{summary.totalLeaveTaken}</span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Weekly Off</span>
                          <span style={styles.summaryValue}>{summary.totalWOCount}</span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Holiday</span>
                          <span style={styles.summaryValue}>{summary.totalHOCount}</span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Total Duration</span>
                          <span style={styles.summaryValue}>{summary.totalDuration}</span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Total T Duration</span>
                          <span style={styles.summaryValue}>{summary.totalTDuration}</span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Over Time</span>
                          <span style={styles.summaryValue}>{summary.totalOverTime}</span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Late By</span>
                          <span style={styles.summaryValue}>{summary.totalLateBy}</span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Early By</span>
                          <span style={styles.summaryValue}>{summary.totalEarlyBy}</span>
                        </div>
                        
                        <div style={styles.summaryItem}>
                          <span style={styles.summaryLabel}>Regular OT</span>
                          <span style={styles.summaryValue}>{summary.totalRegularOT}</span>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// Enhanced Styles
const styles = {
  wrapper: {
    maxWidth: '1400px',
    margin: '80px auto 2rem',
    padding: '1rem',
    backgroundColor: '#f8f9fa',
    borderRadius: 12,
    minHeight: '600px'
  },
  loadingOverlay: {
    position: 'fixed',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: 'rgba(0,0,0,0.5)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 1000
  },
  loadingSpinner: {
    backgroundColor: '#fff',
    padding: '2rem',
    borderRadius: 8,
    fontSize: '18px',
    fontWeight: 'bold'
  },
  header: {
    marginBottom: '2rem'
  },
  title: {
    color: '#343a40',
    marginBottom: '1rem',
    display: 'flex',
    alignItems: 'center',
    fontSize: '28px'
  },
  titleIcon: {
    marginRight: '0.5rem',
    color: '#007bff'
  },
  quickStats: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))',
    gap: '1rem',
    marginBottom: '1rem'
  },
  statCard: {
    backgroundColor: '#fff',
    padding: '1rem',
    borderRadius: 8,
    textAlign: 'center',
    boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
  },
  statNumber: {
    fontSize: '24px',
    fontWeight: 'bold',
    color: '#007bff',
    marginBottom: '0.25rem'
  },
  statLabel: {
    fontSize: '12px',
    color: '#6c757d',
    textTransform: 'uppercase'
  },
  controlsSection: {
    backgroundColor: '#fff',
    padding: '1.5rem',
    borderRadius: 8,
    marginBottom: '1rem',
    boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
  },
  primaryControls: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '1rem',
    alignItems: 'center',
    marginBottom: '1rem'
  },
  searchContainer: {
    position: 'relative',
    flex: '1',
    minWidth: '200px'
  },
  searchIcon: {
    position: 'absolute',
    left: '12px',
    top: '50%',
    transform: 'translateY(-50%)',
    color: '#6c757d'
  },
  searchInput: {
    width: '90%',
    padding: '10px 12px 10px 40px',
    borderRadius: 6,
    border: '1px solid #ced4da',
    fontSize: '14px'
  },
  filterToggle: {
    padding: '10px 16px',
    backgroundColor: '#6c757d',
    color: '#fff',
    border: 'none',
    borderRadius: 6,
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem'
  },
  uploadButton: {
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem',
    backgroundColor: '#28a745',
    color: '#fff',
    padding: '10px 16px',
    borderRadius: 6,
    cursor: 'pointer'
  },
  deleteButton: {
    backgroundColor: '#dc3545',
    color: '#fff',
    padding: '10px 16px',
    borderRadius: 6,
    border: 'none',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem'
  },
  advancedFilters: {
    borderTop: '1px solid #dee2e6',
    paddingTop: '1rem'
  },
  filterRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '1rem',
    marginBottom: '1rem'
  },
  filterGroup: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0.25rem',
    minWidth: '150px',
    flex: '1'
  },
  filterInput: {
    padding: '8px',
    borderRadius: 4,
    border: '1px solid #ced4da',
    fontSize: '14px'
  },
  filterSelect: {
    padding: '8px',
    borderRadius: 4,
    border: '1px solid #ced4da',
    fontSize: '14px'
  },
  clearButton: {
    backgroundColor: '#ffc107',
    color: '#000',
    padding: '8px 12px',
    borderRadius: 4,
    border: 'none',
    cursor: 'pointer',
    marginTop: 'auto'
  },
  uploadInfo: {
    backgroundColor: '#e7f3ff',
    padding: '1rem',
    borderRadius: 6,
    display: 'flex',
    alignItems: 'center',
    gap: '1rem',
    border: '1px solid #b8daff'
  },
  confirmButton: {
    backgroundColor: '#007bff',
    color: '#fff',
    padding: '8px 16px',
    borderRadius: 4,
    border: 'none',
    cursor: 'pointer'
  },
  cancelButton: {
    backgroundColor: '#6c757d',
    color: '#fff',
    padding: '8px 16px',
    borderRadius: 4,
    border: 'none',
    cursor: 'pointer'
  },
  resultsInfo: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '1rem',
    backgroundColor: '#fff',
    borderRadius: 8,
    marginBottom: '1rem',
    boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
  },
  resultsText: {
    color: '#6c757d',
    fontSize: '14px'
  },
  paginationControls: {
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem',
    fontSize: '14px'
  },
  pageSelect: {
    padding: '4px 8px',
    borderRadius: 4,
    border: '1px solid #ced4da'
  },
  tableContainer: {
    backgroundColor: '#fff',
    borderRadius: 8,
    overflow: 'hidden',
    boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
    marginBottom: '1rem'
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse'
  },
  th: {
    backgroundColor: '#007bff',
    color: '#fff',
    padding: '12px 8px',
    textAlign: 'center',
    fontSize: '13px',
    fontWeight: 'bold',
    position: 'sticky',
    top: 0,
    zIndex: 10
  },
  headerContent: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '4px'
  },
  td: {
    padding: '10px 8px',
    textAlign: 'center',
    fontSize: '13px',
    borderBottom: '1px solid #dee2e6'
  },
  row: {
    transition: 'background-color 0.2s',
    cursor: 'pointer'
  },
  summaryButton: {
    backgroundColor: '#007bff',
    border: 'none',
    cursor: 'pointer',
    color: '#fff',
    fontSize: '14px',
    padding: '6px 8px',
    borderRadius: 4,
    transition: 'background-color 0.2s',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },

  noData: {
    textAlign: 'center',
    padding: '3rem',
    color: '#6c757d'
  },
  noDataIcon: {
    fontSize: '48px',
    marginBottom: '1rem',
    opacity: 0.5
  },
  pagination: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    gap: '1rem',
    padding: '1rem',
    backgroundColor: '#fff',
    borderRadius: 8,
    boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
  },
  pageButton: {
    padding: '8px 16px',
    backgroundColor: '#007bff',
    color: 'white',
    border: 'none',
    borderRadius: 4,
    cursor: 'pointer',
    disabled: {
      backgroundColor: '#e9ecef',
      color: '#6c757d',
      cursor: 'not-allowed'
    }
  },
  pageInfo: {
    color: '#6c757d',
    fontWeight: 'bold'
  },
  
  // Popup Styles
  popupOverlay: {
    position: 'fixed',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: 'rgba(0, 0, 0, 0.7)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 1000,
    padding: '20px'
  },
  popupContent: {
    backgroundColor: '#fff',
    borderRadius: 12,
    maxWidth: '900px',
    width: '100%',
    maxHeight: '90vh',
    overflow: 'hidden',
    boxShadow: '0 20px 40px rgba(0, 0, 0, 0.3)',
    display: 'flex',
    flexDirection: 'column'
  },
  popupHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '20px 24px',
    borderBottom: '1px solid #dee2e6',
    backgroundColor: '#f8f9fa'
  },
  popupTitle: {
    margin: 0,
    color: '#343a40',
    display: 'flex',
    alignItems: 'center',
    fontSize: '20px'
  },
  popupIcon: {
    marginRight: '8px',
    color: '#007bff'
  },
  closeButton: {
    backgroundColor: 'transparent',
    border: 'none',
    fontSize: '20px',
    cursor: 'pointer',
    color: '#6c757d',
    padding: '4px',
    borderRadius: 4,
    transition: 'color 0.2s'
  },
  popupBody: {
    padding: '20px 24px',
    overflow: 'auto',
    flex: 1
  },
  noSummaryData: {
    textAlign: 'center',
    padding: '40px 20px',
    color: '#6c757d'
  },
  summaryList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '20px'
  },
  summaryCard: {
    border: '1px solid #dee2e6',
    borderRadius: 8,
    padding: '20px',
    backgroundColor: '#f8f9fa'
  },
  summaryHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '16px',
    paddingBottom: '12px',
    borderBottom: '1px solid #dee2e6'
  },
  summaryBadge: {
    backgroundColor: '#007bff',
    color: '#fff',
    padding: '4px 12px',
    borderRadius: 12,
    fontSize: '12px',
    fontWeight: 'bold'
  },
  summaryGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
    gap: '12px'
  },
  summaryItem: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '8px 12px',
    backgroundColor: '#fff',
    borderRadius: 6,
    border: '1px solid #e9ecef'
  },
  summaryLabel: {
    fontSize: '13px',
    color: '#6c757d',
    fontWeight: '500'
  },
  summaryValue: {
    fontSize: '14px',
    fontWeight: 'bold',
    color: '#343a40'
  }
};

export default EmployeeActivity;
