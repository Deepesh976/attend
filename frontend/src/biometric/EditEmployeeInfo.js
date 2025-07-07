import React, { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

const fieldLabels = [
  { key: 'empStatus', label: 'Status W/L' },
  { key: 'empUnit', label: 'EmpUnit' },
  { key: 'empId', label: 'EmpID' },
  { key: 'empName', label: 'EmpName' },
  { key: 'dob', label: 'Date of Birth' },
  { key: 'bloodGroup', label: 'Blood Group' },
  { key: 'doj', label: 'Date of Joining' },
  { key: 'gender', label: 'Gender' },
  { key: 'qualification', label: 'Academic Qualification' },
  { key: 'experience', label: 'Work Experience' },
  { key: 'personalEmail', label: 'Personal Email id' },
  { key: 'contactNo', label: 'Contact No' },
  { key: 'department', label: 'Department' },
  { key: 'designation', label: 'Designation' },
  { key: 'officialEmail', label: 'Official email id' },
  { key: 'panNo', label: 'PAN No' },
  { key: 'aadharNo', label: 'Aadhar Card No.' },
  { key: 'pfNo', label: 'PF No' },
  { key: 'uanNo', label: 'UAN No' },
  { key: 'esiNo', label: 'ESI No' },
  { key: 'postalAddress', label: 'Postal Address' },
  { key: 'permanentAddress', label: 'Permanent Address' },
  { key: 'bankAccount', label: 'Bank Account Number' },
  { key: 'bankName', label: 'Bank Name' },
  { key: 'ifsc', label: 'IFSC Code' },
  { key: 'bankBranch', label: 'Bank Branch Name' },
  { key: 'fatherName', label: "Father's Name" },
  { key: 'motherName', label: "Mother's Name" },
  { key: 'spouse', label: 'Spouse' },
  { key: 'nomineeName', label: 'Nominee Name' },
  { key: 'emergencyContact', label: 'Emergency Contact no.' },
  { key: 'exitDate', label: 'Exit Date' },
  { key: 'settlementAmount', label: 'Account Settlement (Amount)' },
  { key: 'remarks', label: 'Remarks' },
  { key: 'hiredCtc', label: 'Hired CTC' },
  { key: 'joiningCtc', label: 'CTC in Year:at the time of Joining' },
  { key: 'ctc2025', label: 'CTC in Year:2025' },
  { key: 'yearsWorked', label: 'Total Yrs Worked' },
];

const EditEmployeeInfo = () => {
  const { id } = useParams(); // Mongo _id
  const navigate = useNavigate();
  const [employee, setEmployee] = useState(null);

  useEffect(() => {
    fetch(`http://localhost:5000/api/employee/${id}`)
      .then(res => {
        if (!res.ok) throw new Error('Employee not found');
        return res.json();
      })
      .then(data => setEmployee(data))
      .catch(err => {
        console.error('Fetch Error:', err);
        toast.error('Employee not found');
        navigate('/employee-info');
      });
  }, [id, navigate]);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setEmployee(prev => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    try {
      const res = await fetch(`http://localhost:5000/api/employee/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(employee)
      });

      if (!res.ok) throw new Error('Failed to update');

      toast.success('Employee updated successfully!');
      setTimeout(() => navigate('/employee-info'), 1500);
    } catch (err) {
      console.error('Update error:', err);
      toast.error('Update failed');
    }
  };

  if (!employee) return null;

  return (
    <div style={{ maxWidth: 800, margin: '2rem auto', padding: '2rem', border: '1px solid #ddd', borderRadius: 10, fontFamily: 'Segoe UI' }}>
      <ToastContainer position="top-center" />
      <h2 style={{ textAlign: 'center', marginBottom: '2rem' }}>Edit Employee Info</h2>
      <form onSubmit={handleSubmit} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
        {fieldLabels.map(({ key, label }) => (
          <div key={key} style={{ display: 'flex', flexDirection: 'column' }}>
            <label htmlFor={key} style={{ fontWeight: '600', marginBottom: '0.3rem' }}>{label}</label>
            <input
              type="text"
              id={key}
              name={key}
              value={employee[key] || ''}
              onChange={handleChange}
              style={{ padding: '0.6rem', borderRadius: 6, border: '1px solid #ccc', fontSize: '1rem' }}
            />
          </div>
        ))}
        <div style={{ gridColumn: 'span 2', textAlign: 'center', marginTop: '1.5rem' }}>
          <button type="submit" style={{ backgroundColor: '#007bff', color: '#fff', padding: '0.8rem 2rem', border: 'none', borderRadius: 8, fontSize: '1rem', cursor: 'pointer' }}>
            Save Changes
          </button>
        </div>
      </form>
    </div>
  );
};

export default EditEmployeeInfo;
