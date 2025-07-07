import React, { useState, useRef } from 'react';
import { useNavigate } from 'react-router-dom';
import axios from 'axios';
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

const styles = {
  container: {
    maxWidth: 1200,
    margin: '2rem auto',
    padding: '2rem',
    background: '#fff',
    borderRadius: 10,
    boxShadow: '0 0 15px rgba(0,0,0,0.1)',
    fontFamily: 'Segoe UI, sans-serif',
    boxSizing: 'border-box',
    overflowX: 'hidden'
  },
  heading: {
    fontSize: '1.8rem',
    marginBottom: '1rem',
    textAlign: 'center',
    fontWeight: 600
  },
  sectionTitle: {
    fontSize: '1.2rem',
    fontWeight: 600,
    margin: '2rem 0 1rem',
    borderBottom: '1px solid #ccc',
    paddingBottom: '0.5rem'
  },
  formGroup: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '1rem',
    boxSizing: 'border-box'
  },
  fieldWrapper: {
    position: 'relative',
    flex: '1 1 300px',
    minWidth: '280px',
    boxSizing: 'border-box'
  },
  label: {
    position: 'absolute',
    top: '-10px',
    left: '12px',
    background: '#fff',
    padding: '0 5px',
    fontSize: '0.85rem',
    color: '#555',
    zIndex: 1
  },
  input: {
    width: '100%',
    padding: '0.8rem 2.5rem 0.8rem 0.8rem',
    borderRadius: 8,
    border: '1px solid #ccc',
    fontSize: '1rem',
    outline: 'none',
    boxSizing: 'border-box'
  },
  icon: {
    position: 'absolute',
    right: '10px',
    top: '50%',
    transform: 'translateY(-50%)',
    cursor: 'pointer',
    fontSize: '1.2rem',
    color: '#555'
  },
  greenBtn: {
    backgroundColor: '#28a745',
    color: '#fff',
    padding: '0.6rem 1.5rem',
    border: 'none',
    borderRadius: 8,
    cursor: 'pointer',
    marginTop: '2rem',
    fontWeight: 600
  },
  grayBtn: {
    backgroundColor: '#6c757d',
    color: '#fff',
    padding: '0.6rem 1.5rem',
    border: 'none',
    borderRadius: 8,
    cursor: 'pointer',
    marginLeft: '1rem',
    marginTop: '2rem',
    fontWeight: 600
  }
};

const fieldGroups = {
  'Personal & Employment Details': {
    status_wl: 'Status W/L',
    emp_unit: 'EmpUnit',
    emp_id: 'EmpID',
    emp_name: 'EmpName',
    dob: 'Date of Birth',
    blood_group: 'Blood Group',
    doj: 'Date of Joining',
    gender: 'Gender',
    qualification: 'Academic Qualification',
    experience: 'Work Experience',
    personal_email: 'Personal Email ID',
    contact_no: 'Contact No',
    department: 'Department',
    designation: 'Designation',
    official_email: 'Official Email ID',
    pan_no: 'PAN No',
    aadhar_no: 'Aadhar Card No',
    pf_no: 'PF No',
    uan_no: 'UAN No',
    esi_no: 'ESI No',
    postal_address: 'Postal Address',
    permanent_address: 'Permanent Address',
    bank_account: 'Bank Account Number',
    bank_name: 'Bank Name',
    ifsc: 'IFSC Code',
    bank_branch: 'Bank Branch Name',
    father_name: "Father's Name",
    mother_name: "Mother's Name",
    spouse: 'Spouse',
    nominee_name: 'Nominee Name',
    emergency_contact: 'Emergency Contact no.',
    exit_date: 'Exit Date',
    settlement_amount: 'Account Settlement (Amount)',
    remarks: 'Remarks',
    hired_ctc: 'Hired CTC',
    joining_ctc: 'CTC in Year:at the time of Joining',
    ctc_2025: 'CTC in Year:2025',
    years_worked: 'Total Yrs Worked'
  }
};

const dateFields = ['dob', 'doj', 'exit_date'];

const bloodGroups = ['A+', 'A-', 'B+', 'B-', 'AB+', 'AB-', 'O+', 'O-'];
const genderOptions = ['Male', 'Female', 'Other'];

const AddEmployeeInfo = () => {
  const navigate = useNavigate();

  const [form, setForm] = useState(
    Object.values(fieldGroups).reduce((acc, group) => {
      Object.keys(group).forEach((key) => (acc[key] = ''));
      return acc;
    }, {})
  );

  const dateRefs = {
    dob: useRef(null),
    doj: useRef(null),
    exit_date: useRef(null)
  };

const handleSubmit = async (e) => {
  e.preventDefault();
  try {
    const res = await axios.post('http://localhost:5000/api/employee', form);
    toast.success('Employee Added!');
    setTimeout(() => navigate('/employee-info'), 1000);
  } catch (err) {
    console.error('Error submitting:', err);
    const message = err.response?.data?.message || 'Unknown error occurred';
    toast.error(`Failed to save employee: ${message}`);
  }
};


  return (
    <div style={styles.container}>
      <ToastContainer position="top-center" />
      <h2 style={styles.heading}>Add Employee Info</h2>
      <form onSubmit={handleSubmit}>
        {Object.entries(fieldGroups).map(([section, fields]) => (
          <div key={section}>
            <div style={styles.sectionTitle}>{section}</div>
            <div style={styles.formGroup}>
              {Object.entries(fields).map(([key, label]) => {
                const isDate = dateFields.includes(key);

                // Input types
                const inputType =
                  isDate ? 'date' :
                  key.includes('email') ? 'email' :
                  key === 'nominee_name' ? 'text' :
                  key.includes('contact') || key.includes('no') || key.includes('account') || key.includes('ctc') || key.includes('years') || key.includes('amount')
                    ? 'number'
                    : 'text';

                return (
                  <div key={key} style={styles.fieldWrapper}>
                    <label style={styles.label}>{label}</label>

                    {/* Dropdowns */}
                    {key === 'blood_group' ? (
                      <select
                        style={styles.input}
                        value={form[key]}
                        onChange={(e) => setForm({ ...form, [key]: e.target.value })}
                      >
                        <option value="">Select Blood Group</option>
                        {bloodGroups.map((bg) => (
                          <option key={bg} value={bg}>{bg}</option>
                        ))}
                      </select>
                    ) : key === 'gender' ? (
                      <select
                        style={styles.input}
                        value={form[key]}
                        onChange={(e) => setForm({ ...form, [key]: e.target.value })}
                      >
                        <option value="">Select Gender</option>
                        {genderOptions.map((g) => (
                          <option key={g} value={g}>{g}</option>
                        ))}
                      </select>
                    ) : (
                      <>
                        <input
                          type={inputType}
                          ref={isDate ? dateRefs[key] : null}
                          style={styles.input}
                          placeholder={label}
                          value={form[key]}
                          onClick={() => isDate && dateRefs[key]?.current?.showPicker?.()}
                          onChange={(e) => setForm({ ...form, [key]: e.target.value })}
                        />
                        {isDate && (
                          <span style={styles.icon} onClick={() => dateRefs[key]?.current?.showPicker?.()}>📅</span>
                        )}
                      </>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        ))}
        <button type="submit" style={styles.greenBtn}>Save</button>
        <button type="button" style={styles.grayBtn} onClick={() => navigate('/employee-info')}>Cancel</button>
      </form>
    </div>
  );
};

export default AddEmployeeInfo;
