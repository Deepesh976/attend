import React, { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { toast } from 'react-toastify';
import { jwtDecode } from 'jwt-decode';

const styles = {
  container: {
    maxWidth: '1100px',
    margin: '2rem auto',
    padding: '2rem',
    background: '#f9f9f9',
    borderRadius: '10px',
    boxShadow: '0 0 15px rgba(0,0,0,0.1)',
    fontFamily: 'Segoe UI, Tahoma, Geneva, Verdana, sans-serif',
  },
  title: {
    textAlign: 'center',
    marginBottom: '2rem',
    fontSize: '1.8rem',
    fontWeight: 'bold',
    color: '#333',
  },
  formGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))',
    gap: '1.5rem',
  },
  fieldContainer: {
    display: 'flex',
    flexDirection: 'column',
  },
  label: {
    marginBottom: '0.5rem',
    fontWeight: '600',
    color: '#444',
  },
  input: {
    padding: '0.6rem 0.75rem',
    border: '1px solid #ccc',
    borderRadius: '6px',
    fontSize: '0.95rem',
    outline: 'none',
  },
  buttonContainer: {
    marginTop: '2rem',
    textAlign: 'center',
  },
  submitButton: {
    padding: '0.75rem 2rem',
    backgroundColor: '#007bff',
    color: '#fff',
    border: 'none',
    borderRadius: '6px',
    fontSize: '1rem',
    cursor: 'pointer',
  },
};

const superadminOnlyFields = [
  'actualCTCWithoutLOP', 'consileSalary', 'basic', 'hra', 'cca', 'transportAllowance',
  'otherAllowance1', 'al', 'lop2', 'basic3', 'hra4', 'cca5',
  'transportAllowance6', 'otherAllowance17', 'grossPay', 'plb', 'pf', 'esi',
  'pt', 'tds', 'gpap', 'otherDeductions', 'netPay', 'pfEmployerShare', 'esiEmployerShare', 'bonus'
];

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

const EditEmployeeSalary = () => {
  const navigate = useNavigate();
  const storedData = localStorage.getItem('editEmployee');
  const parsedData = storedData ? JSON.parse(storedData) : null;

  const [formData, setFormData] = useState(parsedData || {});
  const [role, setRole] = useState(null);

  useEffect(() => {
    if (!parsedData || !parsedData._id) {
      toast.error('No employee selected');
      return navigate('/employee-salary-info');
    }

    const token = localStorage.getItem('token');
    if (token) {
      try {
        const decoded = jwtDecode(token);
        setRole(decoded.role);
      } catch (err) {
        console.error('Invalid token');
      }
    }
  }, [parsedData, navigate]);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData((prev) => ({
      ...prev,
      [name]: name === 'dob' || name === 'doj' ? value : value
    }));
  };

  const formatDate = (dateStr) => {
    if (!dateStr) return '';
    const date = new Date(dateStr);
    return !isNaN(date.getTime()) ? date.toISOString().split('T')[0] : '';
  };

  const toNumber = (val) => {
    if (val === null || val === undefined || val === '') return 0;
    return isNaN(Number(val)) ? 0 : Number(val);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();

    try {
      const body = { ...formData };

      // Ensure numeric fields are numbers
      body.lop = toNumber(body.lop);
      body.totalDays = toNumber(body.totalDays);
      body.daysPaid = body.totalDays - body.lop;

      const res = await fetch(`/api/salary/${body._id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(body)
      });

      if (!res.ok) throw new Error('Failed to update salary');

      toast.success('Salary updated successfully!');
      navigate('/employee-salary-info');
    } catch (err) {
      console.error('Update failed', err);
      toast.error('Failed to update salary');
    }
  };

  return (
    <div style={styles.container}>
      <div style={styles.title}>Edit Employee Salary</div>
      <form onSubmit={handleSubmit}>
        <div style={styles.formGrid}>
          {Object.entries(expectedKeys).map(([key, label]) => {
            const value = formData[key] ?? '';
            const isDateField = key === 'dob' || key === 'doj';
            return (
              <div key={key} style={styles.fieldContainer}>
                <label style={styles.label}>{label}</label>
                <input
                  type={isDateField ? 'date' : 'text'}
                  name={key}
                  value={isDateField ? formatDate(value) : value}
                  onChange={handleChange}
                  style={styles.input}
                  disabled={
                    role !== 'superadmin' && superadminOnlyFields.includes(key)
                  }
                />
              </div>
            );
          })}
        </div>
        <div style={styles.buttonContainer}>
          <button type="submit" style={styles.submitButton}>Save</button>
        </div>
      </form>
    </div>
  );
};

export default EditEmployeeSalary;
