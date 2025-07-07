import React, { useState } from 'react';

const styles = {
  container: {
    maxWidth: 1000,
    margin: '2rem auto',
    fontFamily: 'Segoe UI, Tahoma, Geneva, Verdana, sans-serif',
    padding: '2rem',
    background: '#fff',
    borderRadius: 8,
    boxShadow: '0 4px 20px rgba(0,0,0,0.1)',
  },
  heading: {
    textAlign: 'center',
    fontSize: '2rem',
    fontWeight: '700',
    marginBottom: '2rem',
  },
  formRow: {
    display: 'flex',
    justifyContent: 'space-between',
    flexWrap: 'wrap',
    gap: '1rem',
    marginBottom: '2rem',
  },
  inputGroup: {
    flex: '0 0 48%',
    display: 'flex',
    flexDirection: 'column',
    marginBottom: '1rem',
  },
  label: {
    fontWeight: '600',
    marginBottom: '0.5rem',
  },
  input: {
    padding: '0.6rem',
    borderRadius: 6,
    border: '1.5px solid #ccc',
    fontSize: '1rem',
  },
  tableContainer: {
    marginTop: '2rem',
  },
  subHeading: {
    fontSize: '1.3rem',
    fontWeight: '600',
    marginBottom: '1rem',
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
  },
  th: {
    backgroundColor: '#007bff',
    color: '#fff',
    padding: '0.75rem',
    textAlign: 'left',
    border: '1px solid #ccc',
  },
  td: {
    padding: '0.75rem',
    border: '1px solid #ccc',
  },
  salarySection: {
    marginTop: '2rem',
    fontSize: '1.2rem',
    fontWeight: '600',
  },
  submitButton: {
    marginTop: '2rem',
    padding: '0.8rem 2rem',
    fontSize: '1.1rem',
    fontWeight: '700',
    color: '#fff',
    backgroundColor: '#007bff',
    border: 'none',
    borderRadius: 6,
    cursor: 'pointer',
    transition: 'background-color 0.3s ease, box-shadow 0.3s ease',
  },
  submitButtonHover: {
    backgroundColor: '#0056b3',
    boxShadow: '0 4px 15px rgba(0, 86, 179, 0.6)',
  },
};

const toWords = (num) => {
  const a = [
    '', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven',
    'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen',
    'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen',
  ];
  const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  const numberToWords = (n) => {
    if (n === 0) return 'Zero';
    if (n < 20) return a[n];
    if (n < 100) return b[Math.floor(n / 10)] + (n % 10 ? ' ' + a[n % 10] : '');
    if (n < 1000) return a[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' and ' + numberToWords(n % 100) : '');
    if (n < 100000) return numberToWords(Math.floor(n / 1000)) + ' Thousand' + (n % 1000 ? ' ' + numberToWords(n % 1000) : '');
    if (n < 10000000) return numberToWords(Math.floor(n / 100000)) + ' Lakh' + (n % 100000 ? ' ' + numberToWords(n % 100000) : '');
    return numberToWords(Math.floor(n / 10000000)) + ' Crore' + (n % 10000000 ? ' ' + numberToWords(n % 10000000) : '');
  };

  return numberToWords(num) + ' Rupees Only';
};

const GenerateSlip = () => {
  const [earnings, setEarnings] = useState({
    basic: 0,
    hra: 0,
    conveyance: 0,
    transportallowances: 0,
    otherallowances: 0,
    incentives: 0,
  });

  const [deductions, setDeductions] = useState({
    esi: 0,
    pf: 0,
    tax: 0,
    gpap: 0,
    otherdeduction: 0,
    lop: 0,
  });

  const [isHover, setIsHover] = useState(false);

  const totalEarning = Object.values(earnings).reduce((sum, val) => sum + Number(val || 0), 0);
  const totalDeduction = Object.values(deductions).reduce((sum, val) => sum + Number(val || 0), 0);
  const netSalary = totalEarning - totalDeduction;

  const handleChange = (type, field, value) => {
    const parsed = parseFloat(value) || 0;
    if (type === 'earnings') setEarnings({ ...earnings, [field]: parsed });
    else setDeductions({ ...deductions, [field]: parsed });
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    alert(`Submitted! Net Salary is ₹${netSalary}`);
  };

  return (
    <div style={styles.container}>
      <h2 style={styles.heading}>Edit Salary Slip</h2>

      {/* Form Fields */}
      <div style={styles.formRow}>
        {[
          'Name', 'Designation', 'Department', 'Date Of Joining',
          'UAN No', 'ESI No', 'Total Days', 'Days Worked',
          'LOP', 'Annual Leaves', 'PL/ML/BL', 'Bank Account No'
        ].map((label, index) => (
          <div style={styles.inputGroup} key={index}>
            <label style={styles.label}>{label}</label>
            <input
              type="text"
              placeholder=""
              style={styles.input}
            />
          </div>
        ))}
      </div>

      {/* Earnings and Deductions */}
      <div style={styles.tableContainer}>
        <h3 style={styles.subHeading}>Earnings & Deductions</h3>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>Earnings</th>
              <th style={styles.th}>Amount (₹)</th>
              <th style={styles.th}>Deductions</th>
              <th style={styles.th}>Amount (₹)</th>
            </tr>
          </thead>
          <tbody>
            {[
              ['basic', 'esi'],
              ['hra', 'pf'],
              ['conveyance', 'tax'],
              ['transportallowances', 'gpap'],
              ['otherallowances', 'otherdeduction'],
              ['incentives', 'lop']
            ].map(([earn, deduct], i) => (
              <tr key={i}>
                <td style={styles.td}>{earn.charAt(0).toUpperCase() + earn.slice(1).replace(/allowances/, ' Allowances').replace(/other/, 'Other')}</td>
                <td style={styles.td}>
                  <input
                    type="number"
                    style={styles.input}
                    value={earnings[earn] || ''}
                    onChange={(e) => handleChange('earnings', earn, e.target.value)}
                  />
                </td>
                <td style={styles.td}>{deduct.toUpperCase()}</td>
                <td style={styles.td}>
                  <input
                    type="number"
                    style={styles.input}
                    value={deductions[deduct] || ''}
                    onChange={(e) => handleChange('deductions', deduct, e.target.value)}
                  />
                </td>
              </tr>
            ))}
            <tr>
              <td style={styles.td}><strong>Total Earning</strong></td>
              <td style={styles.td}><strong>{totalEarning}</strong></td>
              <td style={styles.td}><strong>Total Deduction</strong></td>
              <td style={styles.td}><strong>{totalDeduction}</strong></td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Net Salary */}
      <div style={styles.salarySection}>
        Net Salary in ₹: {netSalary}
      </div>
      <div style={{ marginTop: '0.5rem', fontStyle: 'italic' }}>
        In Words: {toWords(netSalary)}
      </div>

      {/* Submit Button */}
      <div style={{ textAlign: 'center', marginTop: '2rem' }}>
        <button
          type="submit"
          style={{
            ...styles.submitButton,
            ...(isHover ? styles.submitButtonHover : {}),
          }}
          onMouseEnter={() => setIsHover(true)}
          onMouseLeave={() => setIsHover(false)}
          onClick={handleSubmit}
        >
          Submit
        </button>
      </div>
    </div>
  );
};

export default GenerateSlip;
