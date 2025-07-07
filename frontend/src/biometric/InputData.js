import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import axios from 'axios';
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import { FaUpload, FaTimes, FaTrash, FaSave, FaEdit, FaDownload } from 'react-icons/fa';
import { motion, AnimatePresence } from 'framer-motion';

const styles = {
  container: {
    maxWidth: 1400,
    margin: '2rem auto',
    padding: '2rem',
    background: '#f4f6f8',
    borderRadius: 12,
    fontFamily: 'Segoe UI, sans-serif',
    boxShadow: '0 4px 20px rgba(0,0,0,0.08)'
  },
  topBar: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: '1rem',
    marginTop: '2rem',
    padding: '1rem'
  },
  input: {
    padding: '0.6rem 1rem',
    borderRadius: 8,
    border: '1.5px solid #ccc',
    minWidth: 300,
    fontSize: '1rem',
    outlineColor: '#007bff'
  },
  buttons: {
    display: 'flex',
    gap: '0.75rem',
    flexWrap: 'wrap'
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
  uploadBtn: { backgroundColor: '#17a2b8', color: '#fff' },
  removeBtn: { backgroundColor: '#6c757d', color: '#fff' },
  deleteBtn: { backgroundColor: '#dc3545', color: '#fff' },
  downloadBtn: { backgroundColor: '#17a2b8', color: '#fff'},
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
  editableInput: {
    width: '100%',
    padding: '4px',
    border: '1px solid #ccc',
    borderRadius: 4
  }
};

const expectedKeys = {
  EmpID: 'EmpID',
  EmpName: 'EmpName',
  ActualCTCWithoutLossOfPay: 'Actual CTC Without Loss Of Pay',
  CONSILESALARY: 'CONSILE SALARY',
  Basic: 'Basic',
  HRA: 'HRA',
  CCA: 'CCA',
  TRP_ALW: 'TRP_ALW',
  O_ALW1: 'O_ALW1'
};

const InputData = () => {
  const navigate = useNavigate();
  const [search, setSearch] = useState('');
  const [data, setData] = useState([]);
  const [selected, setSelected] = useState([]);
  const [selectAll, setSelectAll] = useState(false);
  const [editingRow, setEditingRow] = useState(null);
  const [editedRow, setEditedRow] = useState({});
  const role = localStorage.getItem('role');

  useEffect(() => {
    axios.get('http://localhost:5000/api/inputdata')
      .then(res => {
        const resData = Array.isArray(res.data) ? res.data : res.data?.data || [];
        setData(resData);
      })
      .catch(err => toast.error('Failed to fetch data'));
  }, []);

  const handleUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const rawData = XLSX.utils.sheet_to_json(ws, { defval: '' });

      const cleaned = rawData.map((row) => {
        const cleanedRow = {};
        Object.keys(expectedKeys).forEach((key) => {
          const excelKey = expectedKeys[key];
          let value = row[excelKey];
          if (!isNaN(value) && value !== '') value = Math.round(Number(value));
          cleanedRow[key] = value || '';
        });
        return cleanedRow;
      });

      const unique = cleaned.filter(
        (row, index, self) =>
          row.EmpID &&
          index === self.findIndex(r => r.EmpID === row.EmpID)
      );

      const handleDownload = () => {
  if (!filtered.length) {
    toast.info('No data to download');
    return;
  }

  const ws = XLSX.utils.json_to_sheet(filtered.map(({ _id, __v, ...row }) => row));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, 'EmployeeData.xlsx');
};


      try {
        const response = await axios.post('http://localhost:5000/api/inputdata/upload', unique);
        const resData = Array.isArray(response.data) ? response.data : response.data?.data || [];
        setData(resData);
        toast.success('Excel data uploaded and saved');
      } catch (error) {
        toast.error('Upload failed');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleRemoveUpload = async () => {
    try {
      await axios.delete('http://localhost:5000/api/inputdata/clear');
      setData([]);
      setSelected([]);
      setSelectAll(false);
      toast.info('Upload removed and cleared from DB');
    } catch (err) {
      toast.error('Failed to remove upload');
    }
  };

  const handleDeleteSelected = async () => {
    if (selected.length === 0) {
      toast.warn('No rows selected');
      return;
    }

    const confirmDelete = window.confirm('Are you sure to delete selected rows?');
    if (!confirmDelete) return;

    try {
      const idsToDelete = selected.map(i => data[i]?._id).filter(Boolean);
      if (idsToDelete.length === 0) {
        toast.error('No valid rows to delete');
        return;
      }

      await axios.post('http://localhost:5000/api/inputdata/delete-many', { ids: idsToDelete });

      const updated = data.filter((_, i) => !selected.includes(i));
      setData(updated);
      setSelected([]);
      setSelectAll(false);
      toast.success('Selected rows deleted successfully');
    } catch (err) {
      toast.error('Failed to delete selected rows');
    }
  };

  const toggleSelectAll = () => {
    setSelected(selectAll ? [] : data.map((_, i) => i));
    setSelectAll(!selectAll);
  };

  const toggleSelect = (index) => {
    setSelected(prev =>
      prev.includes(index) ? prev.filter(i => i !== index) : [...prev, index]
    );
  };

  const handleEdit = (index) => {
    setEditingRow(index);
    setEditedRow({ ...data[index] });
  };

  const handleEditChange = (key, value) => {
    setEditedRow(prev => ({ ...prev, [key]: value }));
  };

  const handleSave = async (index) => {
    try {
      const updated = { ...editedRow };
      const _id = data[index]._id;
      const res = await axios.put(`http://localhost:5000/api/inputdata/${_id}`, updated);
      const newData = [...data];
      newData[index] = res.data;
      setData(newData);
      toast.success('Row updated');
    } catch (err) {
      toast.error('Update failed');
    }
    setEditingRow(null);
    setEditedRow({});
  };

  const filtered = Array.isArray(data)
    ? data.filter(x => x.EmpName?.toLowerCase().includes(search.toLowerCase()))
    : [];

  return (
    <div style={styles.container}>
      <ToastContainer />
      <div style={styles.topBar}>
        <input
          style={styles.input}
          placeholder="Search by employee name"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
        <div style={styles.buttons}>
          <label style={{ ...styles.btn, ...styles.uploadBtn }}>
            <FaUpload /> Upload
            <input type="file" hidden onChange={handleUpload} accept=".xlsx, .xls" />
          </label>
          <button style={{ ...styles.btn, ...styles.removeBtn }} onClick={handleRemoveUpload}>
            <FaTimes /> Remove Upload
          </button>
          <button style={{ ...styles.btn, ...styles.deleteBtn }} onClick={handleDeleteSelected}>
            <FaTrash /> Delete Selected
          </button>
          <button style={{ ...styles.btn, ...styles.downloadBtn }} onClick={handleDeleteSelected}>
            <FaTrash /> Download
          </button>
        </div>
      </div>

      <div style={styles.tableWrapper}>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>
                <input type="checkbox" checked={selectAll} onChange={toggleSelectAll} />
              </th>
              {Object.keys(expectedKeys).map((key) => (
                <th key={key} style={styles.th}>
                  {expectedKeys[key]}
                </th>
              ))}
              {role === 'superadmin' && <th style={styles.th}>Actions</th>}
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
                  {Object.keys(expectedKeys).map((k) => (
                    <td key={k} style={styles.td}>
                      {editingRow === i ? (
                        <input
                          style={styles.editableInput}
                          value={editedRow[k]}
                          onChange={(e) => handleEditChange(k, e.target.value)}
                        />
                      ) : (
                        row[k]
                      )}
                    </td>
                  ))}
                  {role === 'superadmin' && (
                    <td style={styles.td}>
                      {editingRow === i ? (
                        <button style={styles.btn} onClick={() => handleSave(i)}>
                          <FaSave /> Save
                        </button>
                      ) : (
                        <button style={styles.btn} onClick={() => handleEdit(i)}>
                          <FaEdit /> Edit
                        </button>
                      )}
                    </td>
                  )}
                </motion.tr>
              ))}
            </AnimatePresence>
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default InputData;
