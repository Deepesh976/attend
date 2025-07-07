import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

const styles = {
  container: {
    maxWidth: 900,
    margin: '2rem auto',
    padding: '2rem',
    background: '#ffffff',
    borderRadius: 12,
    boxShadow: '0 8px 20px rgba(0,0,0,0.1)',
    fontFamily: 'Segoe UI, sans-serif',
  },
  heading: {
    textAlign: 'center',
    fontSize: '2rem',
    fontWeight: 'bold',
    marginBottom: '1.5rem',
    color: '#333',
  },
  unitButtons: {
    display: 'flex',
    justifyContent: 'center',
    gap: '1rem',
    marginBottom: '2rem',
  },
  unitBtn: (active) => ({
    padding: '0.6rem 1.2rem',
    borderRadius: 8,
    border: 'none',
    cursor: 'pointer',
    backgroundColor: active ? '#007bff' : '#f1f1f1',
    color: active ? '#fff' : '#333',
    fontWeight: 600,
    fontSize: '1rem',
    boxShadow: '0 2px 6px rgba(0,0,0,0.1)',
  }),
  inputRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '2rem',
    marginBottom: '2rem',
  },
  inputGroup: {
    flex: 1,
    position: 'relative',
  },
  label: {
    fontWeight: 600,
    marginBottom: '0.5rem',
    display: 'block',
    color: '#444',
  },
  input: {
    width: '95%',
    padding: '0.65rem',
    fontSize: '1rem',
    border: '1.5px solid #ccc',
    borderRadius: 6,
    outlineColor: '#007bff',
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    marginTop: '1rem',
  },
  th: {
    backgroundColor: '#007bff',
    color: '#fff',
    padding: '0.75rem',
    border: '1px solid #ccc',
    fontSize: '1rem',
  },
  td: {
    padding: '0.75rem',
    border: '1px solid #ccc',
    textAlign: 'center',
    fontSize: '1rem',
    backgroundColor: '#f9f9f9',
  },
  actionBtn: {
    padding: '0.4rem 0.8rem',
    fontSize: '0.9rem',
    borderRadius: 6,
    border: 'none',
    cursor: 'pointer',
    margin: '0 0.3rem',
  },
  viewBtn: {
    backgroundColor: '#17a2b8',
    color: '#fff',
  },
  downloadBtn: {
    backgroundColor: '#28a745',
    color: '#fff',
  },
  loading: {
    textAlign: 'center',
    fontSize: '1.2rem',
    color: '#666',
    padding: '2rem',
  },
};

const SalarySlip = () => {
  const [nameSearch, setNameSearch] = useState('');
  const [idSearch, setIdSearch] = useState('');
  const [selectedUnit, setSelectedUnit] = useState('');
  const [slips, setSlips] = useState([]);
  const [loading, setLoading] = useState(true);

  // Fetch slips data on component mount
  useEffect(() => {
    fetchSlips();
  }, []);

  const fetchSlips = async () => {
    try {
      setLoading(true);
      const response = await axios.get('/api/slip');
      setSlips(response.data);
    } catch (error) {
      console.error('Error fetching slips:', error);
      toast.error('Failed to fetch salary slips');
    } finally {
      setLoading(false);
    }
  };

  // Get unique units from the slips data
  const availableUnits = [...new Set(slips.map((slip) => slip.empUnit).filter((unit) => unit))];

  const filteredData = slips.filter(
    (slip) =>
      slip.empName.toLowerCase().includes(nameSearch.toLowerCase()) &&
      slip.empId.toLowerCase().includes(idSearch.toLowerCase()) &&
      (selectedUnit === '' || slip.empUnit === selectedUnit)
  );

  const getMonthName = (monthNumber) => {
    const months = [
      'January',
      'February',
      'March',
      'April',
      'May',
      'June',
      'July',
      'August',
      'September',
      'October',
      'November',
      'December',
    ];
    return months[monthNumber - 1] || monthNumber;
  };

  const handleDownload = async (slip) => {
    try {
      const response = await axios.get(`/api/slip/download/${slip._id}`, {
        responseType: 'blob',
      });

      const fileName = `${slip.empName.replace(/\s+/g, '_')}_${getMonthName(slip.month)}_${slip.year}.pdf`;
      const blob = new Blob([response.data], { type: 'application/pdf' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(link.href);

      toast.success('PDF downloaded successfully');
    } catch (error) {
      console.error('Error downloading PDF:', error);
      toast.error('Failed to download PDF');
    }
  };

  const handleView = async (slip) => {
    try {
      const response = await axios.get(`/api/slip/view/${slip._id}`, {
        responseType: 'blob',
      });

      const blob = new Blob([response.data], { type: 'application/pdf' });
      const fileURL = URL.createObjectURL(blob);
      window.open(fileURL, '_blank');

      // Clean up the object URL after a short delay
      setTimeout(() => URL.revokeObjectURL(fileURL), 1000);
    } catch (error) {
      console.error('Error viewing PDF:', error);
      toast.error('Failed to view PDF');
    }
  };

  const handleDelete = async (slipId) => {
    if (!window.confirm('Are you sure you want to delete this salary slip?')) return;

    try {
      await axios.delete(`/api/slip/${slipId}`);
      toast.success('Salary slip deleted successfully');
      // Remove deleted slip from local state
      setSlips((prev) => prev.filter((slip) => slip._id !== slipId));
    } catch (error) {
      console.error('Error deleting slip:', error);
      toast.error('Failed to delete salary slip');
    }
  };

  if (loading) {
    return (
      <div style={styles.container}>
        <div style={styles.loading}>Loading salary slips...</div>
      </div>
    );
  }

  return (
    <div style={styles.container}>
      <h2 style={styles.heading}>Salary Slip Invoice</h2>

      {/* Unit Filter Buttons */}
      <div style={styles.unitButtons}>
        {availableUnits.map((unit) => (
          <button
            key={unit}
            style={styles.unitBtn(selectedUnit === unit)}
            onClick={() => setSelectedUnit((prev) => (prev === unit ? '' : unit))}
          >
            {unit}
          </button>
        ))}
      </div>

      {/* Search Fields */}
      <div style={styles.inputRow}>
        <div style={styles.inputGroup}>
          <label style={styles.label}>Search by Employee Name</label>
          <input
            style={styles.input}
            type="text"
            placeholder="e.g. John Doe"
            value={nameSearch}
            onChange={(e) => setNameSearch(e.target.value)}
          />
        </div>

        <div style={styles.inputGroup}>
          <label style={styles.label}>Search by Employee ID</label>
          <input
            style={styles.input}
            type="text"
            placeholder="e.g. E002"
            value={idSearch}
            onChange={(e) => setIdSearch(e.target.value)}
          />
        </div>
      </div>

      {/* Table */}
      <table style={styles.table}>
        <thead>
          <tr>
            <th style={styles.th}>Unit</th>
            <th style={styles.th}>EmpName</th>
            <th style={styles.th}>EmpID</th>
            <th style={styles.th}>Month</th>
            <th style={styles.th}>Year</th>
            <th style={styles.th}>Actions</th>
          </tr>
        </thead>
        <tbody>
          {filteredData.length > 0 ? (
            filteredData.map((slip) => (
              <tr key={slip._id}>
                <td style={styles.td}>{slip.empUnit || 'N/A'}</td>
                <td style={styles.td}>{slip.empName}</td>
                <td style={styles.td}>{slip.empId}</td>
                <td style={styles.td}>{getMonthName(slip.month)}</td>
                <td style={styles.td}>{slip.year}</td>
                <td style={styles.td}>
                  <button
                    style={{ ...styles.actionBtn, ...styles.viewBtn }}
                    onClick={() => handleView(slip)}
                  >
                    View
                  </button>
                  <button
                    style={{ ...styles.actionBtn, ...styles.downloadBtn }}
                    onClick={() => handleDownload(slip)}
                  >
                    Download
                  </button>
                  <button
                    style={{
                      ...styles.actionBtn,
                      backgroundColor: '#dc3545',
                      color: '#fff',
                      marginLeft: '0.3rem',
                    }}
                    onClick={() => handleDelete(slip._id)}
                  >
                    Delete
                  </button>
                </td>
              </tr>
            ))
          ) : (
            <tr>
              <td style={styles.td} colSpan={6}>
                No matching records
              </td>
            </tr>
          )}
        </tbody>
      </table>

      <ToastContainer />
    </div>
  );
};

export default SalarySlip;
