import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import { FaDownload, FaPlus, FaEdit, FaTrash, FaSync } from 'react-icons/fa';
import { motion, AnimatePresence } from 'framer-motion';

const expectedKeys = {
  year: 'Year',
  month: 'Month',
  empId: 'EmpID',
  empName: 'EmpName',
  department: 'DEPT',
  designation: 'DESIGNATION',
  dob: 'DOB',
  doj: 'DOJ',
  actualCTCWithoutLOP: 'Actual CTC Without Loss Of Pay',
  lopCTC: 'LOP CTC',
  totalDays: 'Total Days',
  daysWorked: 'Days Worked',
  al: 'AL',
  pl: 'PL',
  blOrMl: 'BL/ML',
  lop: 'LOP',
  daysPaid: 'Days Paid',
  consileSalary: 'CONSILE SALARY',
  basic: 'BASIC',
  hra: 'HRA',
  cca: 'CCA',
  transportAllowance: 'TRP_ALW',
  otherAllowance1: 'O_ALW1',
  lop2: 'LOP2',
  basic3: 'BASIC3',
  hra4: 'HRA4',
  cca5: 'CCA5',
  transportAllowance6: 'TRP_ALW6',
  otherAllowance17: 'O_ALW17',
  grossPay: 'Gross Pay',
  plb: 'PLB',
  pf: 'PF',
  esi: 'ESI',
  pt: 'PT',
  tds: 'TDS',
  gpap: 'GPAP',
  otherDeductions: 'OTH_DEDS',
  netPay: 'NET_PAY',
  pfEmployerShare: 'PF Employer Share',
  esiEmployerShare: 'ESI Employer Share',
  bonus: 'Bonus'
};

const totalKeys = Object.keys(expectedKeys).filter(
  key => !['year', 'month', 'empId', 'empName', 'department', 'designation', 'dob', 'doj'].includes(key)
);

const styles = {
  container: {
    maxWidth: 1400,
    margin: '5rem auto',
    padding: '2rem',
    background: '#f4f6f8',
    borderRadius: 12,
    fontFamily: 'Segoe UI, sans-serif',
    boxShadow: '0 4px 20px rgba(0,0,0,0.08)'
  },
  toolbar: {
    display: 'flex',
    flexWrap: 'wrap',
    justifyContent: 'space-between',
    marginBottom: '1.5rem',
    alignItems: 'center'
  },
  searchGroup: {
    display: 'flex',
    gap: '1rem',
    alignItems: 'center'
  },
  input: {
    padding: '0.6rem 1rem',
    borderRadius: 8,
    border: '1.5px solid #ccc',
    minWidth: 200,
    fontSize: '1rem',
    outlineColor: '#007bff'
  },
  select: {
    padding: '0.6rem 1rem',
    borderRadius: 8,
    border: '1.5px solid #ccc',
    fontSize: '1rem',
    outlineColor: '#007bff'
  },
  btn: {
    padding: '0.55rem 1.2rem',
    fontSize: '0.95rem',
    borderRadius: 6,
    border: 'none',
    cursor: 'pointer',
    fontWeight: 600,
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem'
  },
  editBtn: { backgroundColor: '#ffc107', color: '#222' },
  deleteBtn: { backgroundColor: '#dc3545', color: '#fff' },
  greenBtn: { backgroundColor: '#17a2b8', color: '#fff' },
  purpleBtn: { backgroundColor: '#6f42c1', color: '#fff' },
  tableWrapper: { overflowX: 'auto' },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    marginTop: '1rem',
    backgroundColor: '#fff'
  },
  th: {
    backgroundColor: '#007bff',
    color: '#fff',
    padding: '0.5rem 0.75rem',
    border: '1px solid #ddd',
    fontWeight: 700,
    textAlign: 'center',
    whiteSpace: 'nowrap'
  },
  td: {
    padding: '0.4rem 0.6rem',
    border: '1px solid #eee',
    textAlign: 'center',
    fontSize: '0.9rem',
    color: '#333',
    whiteSpace: 'nowrap'
  },
  rowHighlight: {
    backgroundColor: '#e8f4fd'
  },
  totalRow: {
    backgroundColor: '#e0e0e0',
    fontWeight: 'bold'
  },
  downloadBtn: {
    marginTop: '1.5rem',
    backgroundColor: '#007bff',
    color: '#fff',
    padding: '0.6rem 1.2rem',
    borderRadius: 6,
    border: 'none',
    fontSize: '1rem',
    cursor: 'pointer'
  }
};

const EmployeeSalaryInfo = () => {
  const navigate = useNavigate();
  const [search, setSearch] = useState('');
  const [monthFilter, setMonthFilter] = useState('');
  const [yearFilter, setYearFilter] = useState('');
  const [data, setData] = useState([]);
  const [selected, setSelected] = useState([]);
  const [selectAll, setSelectAll] = useState(false);

  const fetchSalaryData = () => {
    fetch('/api/salary')
      .then(res => res.json())
      .then(setData)
      .catch(() => toast.error('Failed to load salary data.'));
  };

  useEffect(() => {
    fetchSalaryData();
  }, []);

  const filtered = data.filter(item => {
    const val = search.toLowerCase();
    const matchText =
      item.empName?.toLowerCase().includes(val) ||
      item.empId?.toLowerCase().includes(val) ||
      item.department?.toLowerCase().includes(val);
    const matchYear = !yearFilter || item.year === yearFilter;
    const matchMonth = !monthFilter || item.month === monthFilter;
    return matchText && matchYear && matchMonth;
  });

  const toggleSelectAll = () => {
    setSelectAll(!selectAll);
    setSelected(!selectAll ? data.map((_, i) => i) : []);
  };

  const toggleSelect = index => {
    setSelected(prev =>
      prev.includes(index) ? prev.filter(i => i !== index) : [...prev, index]
    );
  };

  const handleEdit = () => {
    if (selected.length !== 1) return toast.warn('Select exactly one employee to edit');
    const empData = data[selected[0]];
    localStorage.setItem('editEmployee', JSON.stringify(empData));
    navigate(`/edit-employee-salary-info/${empData._id}`);
  };

  const handleDelete = () => {
    if (selected.length === 0) return toast.warn('No employee selected');
    if (!window.confirm(`Delete ${selected.length} employee(s)?`)) return;
    selected.forEach(i => {
      const id = data[i]?._id;
      if (id) fetch(`/api/salary/${id}`, { method: 'DELETE' }).then(fetchSalaryData);
    });
    toast.success('Selected employee(s) deleted!');
    setSelected([]);
    setSelectAll(false);
  };

  const handleRegenerate = () => {
    fetch('/api/salary/generate-from-employee', { method: 'POST' })
      .then(() => {
        toast.success('Salary regenerated!');
        setTimeout(fetchSalaryData, 1000);
      })
      .catch(() => toast.error('Failed to regenerate salary'));
  };

  const handleDownload = () => {
    const exportData = data.map(row => {
      const newRow = {};
      Object.keys(expectedKeys).forEach(key => {
        const val = row[key];
        newRow[expectedKeys[key]] = totalKeys.includes(key)
          ? isNaN(val) ? '' : Math.round(parseFloat(val))
          : val;
      });
      return newRow;
    });
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Salary Info');
    XLSX.writeFile(wb, 'salary_info.xlsx');
  };

  return (
    <div style={styles.container}>
      <ToastContainer />
      <div style={styles.toolbar}>
        <div style={styles.searchGroup}>
          <input
            style={styles.input}
            placeholder="Search by Emp Name, ID, Dept"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
          <select style={styles.select} value={yearFilter} onChange={e => setYearFilter(e.target.value)}>
            <option value="">All Years</option>
            {[...new Set(data.map(d => d.year))].map(y => <option key={y} value={y}>{y}</option>)}
          </select>
          <select style={styles.select} value={monthFilter} onChange={e => setMonthFilter(e.target.value)}>
            <option value="">All Months</option>
            {[...new Set(data.map(d => d.month))].map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
        <div style={{ display: 'flex', gap: '0.75rem' }}>
          <button style={{ ...styles.btn, ...styles.editBtn }} onClick={handleEdit}><FaEdit /> Edit</button>
          <button style={{ ...styles.btn, ...styles.deleteBtn }} onClick={handleDelete}><FaTrash /> Delete</button>
          <button style={{ ...styles.btn, ...styles.greenBtn }} onClick={() => navigate('/input-data')}><FaPlus /> Insert</button>
          <button style={{ ...styles.btn, ...styles.purpleBtn }} onClick={handleRegenerate}><FaSync /> Regenerate</button>
        </div>
      </div>

      <div style={styles.tableWrapper}>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>
                <input type="checkbox" checked={selectAll} onChange={toggleSelectAll} />
              </th>
              {Object.values(expectedKeys).map(h => (
                <th key={h} style={styles.th}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            <AnimatePresence>
              {filtered.map((row, i) => (
                <motion.tr
                  key={i}
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -10 }}
                  style={selected.includes(i) ? styles.rowHighlight : {}}
                >
                  <td style={styles.td}>
                    <input
                      type="checkbox"
                      checked={selected.includes(i)}
                      onChange={() => toggleSelect(i)}
                    />
                  </td>
                  {Object.keys(expectedKeys).map(k => (
                    <td key={k} style={styles.td}>
                      {totalKeys.includes(k) && !isNaN(row[k])
                        ? Math.round(parseFloat(row[k]))
                        : row[k]}
                    </td>
                  ))}
                </motion.tr>
              ))}
            </AnimatePresence>

<tr style={styles.totalRow}>
  <td style={styles.td}><strong>Total</strong></td> {/* "Total" label in first cell */}
  {Object.keys(expectedKeys).map((key, idx) => {
    if (!totalKeys.includes(key)) {
      return <td key={key} style={styles.td}></td>;
    }
    const sum = filtered.reduce((acc, row) => acc + (parseFloat(row[key]) || 0), 0);
    return <td key={key} style={styles.td}><strong>{Math.round(sum)}</strong></td>;
  })}
</tr>

          </tbody>
        </table>
      </div>

      <button style={styles.downloadBtn} onClick={handleDownload}>
        <FaDownload /> Download Excel
      </button>
    </div>
  );
};

export default EmployeeSalaryInfo;
