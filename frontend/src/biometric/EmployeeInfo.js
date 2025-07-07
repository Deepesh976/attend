import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { useNavigate } from 'react-router-dom';
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
// import { saveAs } from 'file-saver';

const styles = {
  container: { maxWidth: 1280, margin: '2rem auto', padding: '2rem', background: '#f9f9f9', borderRadius: 12, fontFamily: 'Segoe UI, sans-serif', boxShadow: '0 4px 20px rgba(0,0,0,0.08)' },
  heading: { textAlign: 'center', marginBottom: '2rem', fontWeight: 700, fontSize: '2rem', color: '#222' },
  filters: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem', marginBottom: '1.5rem' },
  leftFilters: { display: 'flex', alignItems: 'center', gap: '1rem' },
  rightButtons: { display: 'flex', alignItems: 'center', gap: '1rem' },
  input: { padding: '0.6rem 1rem', borderRadius: 8, border: '1.5px solid #ccc', minWidth: 220, fontSize: '1rem', outlineColor: '#007bff' },
  uploadLabel: { backgroundColor: '#17a2b8', color: '#fff', padding: '0.6rem 1.2rem', borderRadius: 8, cursor: 'pointer', fontWeight: 600 },
  hiddenInput: { display: 'none' },
  greenBtn: { backgroundColor: '#28a745', color: '#fff', padding: '0.6rem 1.2rem', border: 'none', borderRadius: 8, cursor: 'pointer', fontWeight: 600 },
  yellowBtn: { backgroundColor: '#ffc107', color: '#000', padding: '0.6rem 1.2rem', border: 'none', borderRadius: 8, cursor: 'pointer', fontWeight: 600 },
  redBtn: { backgroundColor: '#dc3545', color: '#fff', padding: '0.6rem 1.2rem', border: 'none', borderRadius: 8, cursor: 'pointer', fontWeight: 600 },
  blueBtn: { backgroundColor: '#007bff', color: '#fff', padding: '0.6rem 1.2rem', border: 'none', borderRadius: 8, cursor: 'pointer', fontWeight: 600 },
  tableWrapper: { overflowX: 'auto' },
  table: { width: '100%', borderCollapse: 'collapse', marginTop: '1rem', backgroundColor: '#fff' },
  th: { backgroundColor: '#007bff', color: '#fff', padding: '0.8rem', border: '1px solid #ddd', fontWeight: 700 },
  td: { padding: '0.75rem', border: '1px solid #eee', textAlign: 'center', fontSize: '0.95rem', color: '#333' },
  pagination: { marginTop: '2rem', display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '1rem' },
  button: { padding: '0.5rem 1.2rem', border: 'none', borderRadius: 8, fontWeight: 600, fontSize: '1rem', color: '#fff', cursor: 'pointer' },
  blueButton: { backgroundColor: '#007bff' },
  blueButtonDisabled: { backgroundColor: '#a6c8ff', cursor: 'not-allowed' },
  checkbox: { cursor: 'pointer', width: 18, height: 18 }
};

const expectedKeys = {
  empStatus: 'Status W/ L',
  empUnit: 'EmpUnit',
  empId: 'EmpId',
  empName: 'EmpName',
  dob: 'Date of Birth',
  bloodGroup: 'Blood Group',
  doj: 'Date of Joining',
  gender: 'Gender',
  qualification: 'Academic Qualification',
  experience: 'Work Experience',
  personalEmail: 'Personal Email id',
  contactNo: 'Contact No',
  department: 'Department',
  designation: 'Designation',
  officialEmail: 'Official email id',
  panNo: 'PAN No',
  aadharNo: 'Aadhar Card No.',
  pfNo: 'PF No',
  uanNo: 'UAN No',
  esiNo: 'ESI No',
  postalAddress: 'Postal Address',
  permanentAddress: 'Permanent Address',
  bankAccount: 'Bank Account Number',
  bankName: 'Bank Name',
  ifsc: 'IFSC Code',
  bankBranch: 'Bank Branch Name',
  fatherName: "Father's Name",
  motherName: "Mother's Name",
  spouse: 'Spouse',
  nomineeName: 'Nominee Name',
  emergencyContact: 'Emergency Contact no.',
  exitDate: 'Exit Date',
  settlementAmount: 'Account Settlement (Amount)',
  remarks: 'Remarks',
  hiredCtc: 'Hired CTC',
  joiningCtc: 'CTC in Year:at the time of Joining',
  ctc2025: 'CTC in Year:2025',
  yearsWorked: 'Total Yrs Worked'
};

const EmployeeInfo = () => {
  const navigate = useNavigate();
  const [data, setData] = useState([]);
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [selectedIds, setSelectedIds] = useState([]);
  const [fileUploaded, setFileUploaded] = useState(false);
  const [loading, setLoading] = useState(true);
  const limit = 10;

  const fetchEmployeeData = () => {
    setLoading(true);
    
    fetch('http://localhost:5000/api/employee')
      .then(res => {
        if (!res.ok) throw new Error('Failed to fetch employees');
        return res.json();
      })
      .then(data => setData(data))
      .catch(err => {
        console.error('Fetch error:', err);
        toast.error('Could not load employee data');
      })
      .finally(() => setLoading(false));
  };

  useEffect(() => {
    fetchEmployeeData();
  }, []);

const handleUpload = async (e) => {
  e.preventDefault();

  const file = e?.target?.files?.[0];
  if (!file || file.size === 0) {
    toast.error('⚠️ No file selected or file is empty');
    return;
  }

  if (fileUploaded) {
    toast.warn('⚠️ File already uploaded');
    return;
  }


  const reader = new FileReader();
  reader.onload = async (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1 });

    const headers = rawData[0];
    const rows = rawData.slice(1);

    const jsonData = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    // ✅ Check for empty or invalid Excel
    if (!jsonData || jsonData.length === 0) {
      toast.error('❌ Excel file appears empty or invalid');
      return;
    }

    const normalize = str => str.toLowerCase().replace(/[\s_:/.-]/g, '').replace(/[^\w]/g, '');

    const parseDate = (value) => {
      if (!value || typeof value !== 'string') return null;
      const parts = value.split(/[./-]/);
      if (parts.length !== 3) return null;
      const [day, month, year] = parts.map(Number);
      if (!day || !month || !year) return null;
      return new Date(year, month - 1, day);
    };
    
    const parseNumber = (val) => {
      if (val === null || val === undefined || val === '') return '';
      const num = parseFloat(val);
    return isNaN(num) ? '' : num;
    };

    const mappedData = jsonData.map(row => {
      const mappedRow = {};
      for (const [key, label] of Object.entries(expectedKeys)) {
        const matched = Object.keys(row).find(col => normalize(col) === normalize(label));
        let value = matched ? row[matched] : '';

        if (['dob', 'doj', 'exitDate'].includes(key)) {
          value = parseDate(value);
        } else if (['settlementAmount', 'hiredCtc', 'joiningCtc', 'ctc2025', 'yearsWorked'].includes(key)) {
          value = parseNumber(value);
        } else if (key === 'empStatus') {
          value = String(value).toUpperCase();
        }
        if (key === 'empId') {
          value = String(value).trim();
        }
        mappedRow[key] = value;
      }
      return mappedRow;
    });

    if (mappedData.length === 0) {
      toast.error('❌ No valid employee data to upload');
      return;
    }

    const invalidRows = mappedData.filter(emp => !emp.empId || !emp.empName);
    if (invalidRows.length > 0) {
      toast.error(`❌ ${invalidRows.length} row(s) missing empId or empName.`);
      return;
    }

    console.log("🚀 mappedData to send:", mappedData);

    try {
      const res = await fetch('http://localhost:5000/api/employee/upload', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ employees: mappedData })
      });

      if (!res.ok) throw new Error('Upload failed');

      toast.success(`✅ Uploaded ${mappedData.length} rows`);
      fetchEmployeeData();
      setFileUploaded(true);
    } catch (err) {
      console.error(err);
      toast.error('❌ Failed to upload');
    }
  };

  reader.readAsArrayBuffer(file);
};

const formatDate = (dateString) => {
  if (!dateString) return '';
  return new Date(dateString).toISOString().split('T')[0];
};

const handleEdit = () => {
  if (selectedIds.length === 0) {
    toast.warn('Please select an employee to edit.');
    return;
  }
  if (selectedIds.length > 1) {
    toast.warn('You can only edit one employee at a time.');
    return;
  }

  const selectedEmployee = data.find(emp => emp.empId === selectedIds[0]);
  if (selectedEmployee && selectedEmployee._id) {
    navigate(`/edit-employee-info/${selectedEmployee._id}`);
  } else {
    toast.error('Could not find selected employee’s ID');
  }
};


  const handleDelete = () => {
    if (selectedIds.length === 0) return toast.warn('Select at least one row to delete.');
    fetch('http://localhost:5000/api/employee/delete', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ids: selectedIds })
    })
      .then(res => res.json())
      .then(() => {
        toast.success('Deleted successfully');
        setSelectedIds([]);
        fetchEmployeeData();
      })
      .catch(err => {
        console.error('Delete failed:', err);
        toast.error('Failed to delete');
      });
  };

const filteredData = data.filter(emp =>
  emp.empName?.toLowerCase().includes(search.toLowerCase()) ||
  emp.empId?.toLowerCase().includes(search.toLowerCase())
);


  const totalPages = Math.ceil(filteredData.length / limit);
  const paginatedData = filteredData.slice((page - 1) * limit, page * limit);

  useEffect(() => { setPage(1); }, [search]);

  const toggleSelect = (id) => {
    setSelectedIds(prev =>
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

return (
  <div style={styles.container}>
    <ToastContainer position="top-center" />
    <h2 style={styles.heading}>Employee Information</h2>

 <div style={styles.filters}>
  <div style={styles.leftFilters}>
    <label htmlFor="upload" style={styles.uploadLabel}>Upload File</label>
    <input id="upload" type="file" accept=".xlsx, .xls" onChange={handleUpload} style={styles.hiddenInput} />

    <button
      onClick={async () => {
        try {
          const res = await fetch('http://localhost:5000/api/employee/all', { method: 'DELETE' });
          if (!res.ok) throw new Error('Failed to delete all records');
          toast.success('🗑️ All employee data deleted');
          setData([]);
          setFileUploaded(false);
        } catch (err) {
          console.error('❌ Error deleting all employees:', err);
          toast.error('Failed to delete all employee data');
        }
      }}
      style={styles.redBtn}
    >
      Remove File
    </button>

    {fileUploaded && (
      <span style={{ color: '#28a745', fontWeight: 600 }}>File Uploaded ✓</span>
    )}
  </div>

  <div style={styles.rightButtons}>
    <input
      style={styles.input}
      type="text"
      placeholder="Search by Name or ID"
      value={search}
      onChange={e => setSearch(e.target.value)}
    />
    <button onClick={handleEdit} style={styles.yellowBtn}>Edit</button>
    <button onClick={handleDelete} style={styles.redBtn}>Delete</button>
  </div>
</div>

    <div style={styles.tableWrapper}>
      <table style={styles.table}>
        <thead>
          <tr>
            <th style={styles.th}></th>
            {Object.values(expectedKeys).map((label, idx) => (
              <th key={idx} style={styles.th}>{label}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {loading ? (
            <tr>
              <td colSpan={Object.keys(expectedKeys).length + 1} style={styles.td}>Loading...</td>
            </tr>
          ) : paginatedData.length === 0 ? (
            <tr>
              <td colSpan={Object.keys(expectedKeys).length + 1} style={styles.td}>No records found.</td>
            </tr>
          ) : (
            paginatedData.map((emp, i) => (
              <tr key={i} style={selectedIds.includes(emp.empId) ? { backgroundColor: '#e6f7ff' } : {}}>
                <td style={styles.td}>
                  <input
                    type="checkbox"
                    style={styles.checkbox}
                    checked={selectedIds.includes(emp.empId)}
                    onChange={() => toggleSelect(emp.empId)}
                  />
                </td>
                {Object.keys(expectedKeys).map((key, idx) => (
                  <td key={idx} style={styles.td}>
                    {['dob', 'doj', 'exitDate'].includes(key) && emp[key]
                      ? formatDate(emp[key])
                      : emp[key]}
                  </td>
                ))}
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>

    <div style={styles.pagination}>
      <button
        style={{
          ...styles.button,
          ...styles.blueButton,
          ...(page === 1 && styles.blueButtonDisabled)
        }}
        onClick={() => setPage(p => p - 1)}
        disabled={page === 1}
      >
        Previous
      </button>
      <span>Page {page} of {totalPages}</span>
      <button
        style={{
          ...styles.button,
          ...styles.blueButton,
          ...(page === totalPages && styles.blueButtonDisabled)
        }}
        onClick={() => setPage(p => p + 1)}
        disabled={page === totalPages}
      >
        Next
      </button>
    </div>
  </div>
);

}

export default EmployeeInfo;
