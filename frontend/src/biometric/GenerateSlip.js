import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

const GenerateSlip = () => {
  const [employees, setEmployees] = useState([]);
  const [selectedEmpId, setSelectedEmpId] = useState('');
  const [loading, setLoading] = useState(false);

  const [formData, setFormData] = useState({
    empId: '',
    empName: '',
    empUnit: '',
    designation: '',
    department: '',
    dateOfJoining: '',
    uanNo: '',
    esiNo: '',
    bankAccountNo: '',
    totalDays: '',
    daysWorked: '',
    lop: '',
    annualLeaves: '',
    plMlBl: '',
    earnings: {
      basic: '',
      hra: '',
      conveyance: '',
      transportAllowances: '',
      otherAllowances: '',
      incentives: ''
    },
    deductions: {
      esi: '',
      pf: '',
      tax: '',
      gpap: '',
      otherDeductions: '',
      lop: ''
    }
  });

  useEffect(() => {
    const fetchEmployees = async () => {
      try {
        const res = await axios.get('/api/employee');
        setEmployees(res.data);
      } catch {
        toast.error('Failed to load employees');
      }
    };
    fetchEmployees();
  }, []);

  useEffect(() => {
    if (!selectedEmpId) return;
    const fetchDetails = async () => {
      try {
        const res = await axios.get(`/api/slip/details/${selectedEmpId}`);
        const { employee, salary, earnings, deductions, annualLeaves, plMlBl } = res.data;

        setFormData({
          empId: selectedEmpId,
          empName: employee.empName || '',
          empUnit: employee.empUnit || '',
          designation: employee.designation || '',
          department: employee.department || '',
          dateOfJoining: employee.doj ? employee.doj.split('T')[0] : '',
          uanNo: employee.uanNo || '',
          esiNo: employee.esiNo || '',
          bankAccountNo: employee.bankAccount || '',
          totalDays: round(salary.totalDays),
          daysWorked: salary.daysWorked,
          lop: salary.lop,
          annualLeaves: annualLeaves,
          plMlBl: plMlBl,
          earnings: {
            basic: round(earnings.basic),
            hra: round(earnings.hra),
            conveyance: round(earnings.conveyance),
            transportAllowances: round(earnings.transportAllowances),
            otherAllowances: round(earnings.otherAllowances),
            incentives: round(earnings.incentives)
          },
          deductions: {
            esi: round(deductions.esi),
            pf: round(deductions.pf),
            tax: round(deductions.tax),
            gpap: round(deductions.gpap),
            otherDeductions: round(deductions.otherDeductions),
            lop: round(deductions.lop)
          }
        });
      } catch {
        toast.error('Failed to fetch employee details');
      }
    };
    fetchDetails();
  }, [selectedEmpId]);

  const round = (value) => {
    const n = parseFloat(value);
    return isNaN(n) ? value : Math.round(n);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    const numericFields = ['daysWorked', 'lop', 'annualLeaves', 'plMlBl'];
    const shouldRound = !numericFields.includes(name);
    const updatedValue = shouldRound ? round(value) : value;

    if (name.startsWith('earnings.')) {
      const key = name.split('.')[1];
      setFormData((prev) => ({
        ...prev,
        earnings: { ...prev.earnings, [key]: updatedValue }
      }));
    } else if (name.startsWith('deductions.')) {
      const key = name.split('.')[1];
      setFormData((prev) => ({
        ...prev,
        deductions: { ...prev.deductions, [key]: updatedValue }
      }));
    } else {
      setFormData((prev) => ({
        ...prev,
        [name]: updatedValue
      }));
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    try {
      const res = await axios.post('/api/slip', formData);
      if (res.data.pdfUrl) {
        window.open(`http://localhost:5000${res.data.pdfUrl}`, '_blank');
        toast.success('Salary slip generated successfully!');
      }
      setSelectedEmpId('');
    } catch (error) {
      toast.error(error.response?.data?.error || 'Error generating slip');
    } finally {
      setLoading(false);
    }
  };

  const styles = {
    container: {
      maxWidth: '1100px',
      margin: '2rem auto',
      padding: '2rem',
      backgroundColor: '#fff',
      borderRadius: '12px',
      boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
      fontFamily: 'Segoe UI, sans-serif'
    },
    heading: {
      fontSize: '2rem',
      textAlign: 'center',
      marginBottom: '2rem'
    },
    sectionTitle: {
      fontSize: '1.3rem',
      margin: '2rem 0 1rem 0',
      fontWeight: 'bold',
      borderBottom: '1px solid #ccc',
      paddingBottom: '0.5rem'
    },
    grid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))',
      gap: '1rem',
      marginBottom: '1.5rem'
    },
    label: {
      fontWeight: '500',
      marginBottom: '0.2rem',
      display: 'block'
    },
    input: {
      width: '90%',
      padding: '0.5rem',
      borderRadius: '6px',
      border: '1px solid #ccc',
      fontSize: '1rem'
    },
    button: {
      padding: '0.75rem 2rem',
      backgroundColor: '#007bff',
      color: '#fff',
      border: 'none',
      borderRadius: '6px',
      fontSize: '1rem',
      cursor: 'pointer',
      marginTop: '1.5rem'
    }
  };

  const toTitleCase = (str) => str.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase());

  return (
    <div style={styles.container}>
      <h1 style={styles.heading}>Salary Slip Generation</h1>
      <form onSubmit={handleSubmit}>
        <div>
          <label style={styles.label}>Select Employee</label>
          <select
            style={styles.input}
            value={selectedEmpId}
            onChange={(e) => setSelectedEmpId(e.target.value)}
            required
          >
            <option value="">-- Select Employee --</option>
            {employees.map((emp) => (
              <option key={emp._id} value={emp.empId}>
                {emp.empName} ({emp.empId})
              </option>
            ))}
          </select>
        </div>

        <h2 style={styles.sectionTitle}>Employee Information</h2>
        <div style={styles.grid}>
          {[
            ['Employee Name', 'empName'],
            ['Designation', 'designation'],
            ['Department', 'department'],
            ['Date of Joining', 'dateOfJoining', 'date'],
            ['UAN No', 'uanNo'],
            ['ESI No', 'esiNo'],
            ['Bank Account No', 'bankAccountNo']
          ].map(([label, name, type = 'text']) => (
            <div key={name}>
              <label style={styles.label}>{label}</label>
              <input
                type={type}
                name={name}
                value={formData[name]}
                onChange={handleChange}
                style={styles.input}
              />
            </div>
          ))}
        </div>

        <h2 style={styles.sectionTitle}>Attendance & Leave</h2>
        <div style={styles.grid}>
          {[
            ['Total Days', 'totalDays'],
            ['Days Worked', 'daysWorked'],
            ['Loss of Pay', 'lop'],
            ['Annual Leaves', 'annualLeaves'],
            ['PL / ML / BL', 'plMlBl']
          ].map(([label, name]) => (
            <div key={name}>
              <label style={styles.label}>{label}</label>
              <input
                type="number"
                name={name}
                value={formData[name]}
                onChange={handleChange}
                style={styles.input}
              />
            </div>
          ))}
        </div>

        <h2 style={styles.sectionTitle}>Earnings</h2>
        <div style={styles.grid}>
          {Object.entries(formData.earnings).map(([key, val]) => (
            <div key={key}>
              <label style={styles.label}>{toTitleCase(key)}</label>
              <input
                type="number"
                name={`earnings.${key}`}
                value={val}
                onChange={handleChange}
                style={styles.input}
              />
            </div>
          ))}
        </div>

        <h2 style={styles.sectionTitle}>Deductions</h2>
        <div style={styles.grid}>
          {Object.entries(formData.deductions).map(([key, val]) => (
            <div key={key}>
              <label style={styles.label}>{toTitleCase(key)}</label>
              <input
                type="number"
                name={`deductions.${key}`}
                value={val}
                onChange={handleChange}
                style={styles.input}
              />
            </div>
          ))}
        </div>

        <button type="submit" style={styles.button} disabled={loading}>
          {loading ? 'Generating...' : 'Generate Slip'}
        </button>
      </form>
      <ToastContainer />
    </div>
  );
};

export default GenerateSlip;
