import React, { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

const fieldLabels = [
  { key: 'empStatus', label: 'Emp Status' },
  { key: 'empUnit', label: 'Emp Unit' },
  { key: 'empId', label: 'Emp Id' },
  { key: 'empName', label: 'Emp Name' },
  { key: 'dob', label: 'DOB' },
  { key: 'bloodGroup', label: 'Blood Group' },
  { key: 'doj', label: 'DOJ' },
  { key: 'gender', label: 'Gender' },
  { key: 'qualification', label: 'Qualification' },
  { key: 'experience', label: 'Experience' },
  { key: 'personalEmail', label: 'Personal Email' },
  { key: 'contactNo', label: 'Contact No' },
  { key: 'department', label: 'Department' },
  { key: 'designation', label: 'Designation' },
  { key: 'officialEmail', label: 'Official Email' },
  { key: 'panNo', label: 'PAN No' },
  { key: 'aadharNo', label: 'Aadhar No' },
  { key: 'pfNo', label: 'PF No' },
  { key: 'uanNo', label: 'UAN No' },
  { key: 'esiNo', label: 'ESI No' },
  { key: 'postalAddress', label: 'Postal Address' },
  { key: 'permanentAddress', label: 'Permanent Address' },
  { key: 'bankAccount', label: 'Bank Account' },
  { key: 'bankName', label: 'Bank Name' },
  { key: 'ifsc', label: 'IFSC' },
  { key: 'bankBranch', label: 'Bank Branch' },
  { key: 'fatherName', label: 'Father Name' },
  { key: 'motherName', label: 'Mother Name' },
  { key: 'spouse', label: 'Spouse' },
  { key: 'nomineeName', label: 'Nominee Name' },
  { key: 'emergencyContact', label: 'Emergency Contact' },
  { key: 'exitDate', label: 'Exit Date' },
  { key: 'settlementAmount', label: 'Settlement Amount' },
  { key: 'remarks', label: 'Remarks' },
  { key: 'hiredCtc', label: 'Hired CTC' },
  { key: 'joiningCtc', label: 'Joining CTC' },
  { key: 'ctc2025', label: 'CTC 2025' },
  { key: 'yearsWorked', label: 'Years Worked' },
];

const dateFields = ['dob', 'doj', 'exitDate'];

const EditEmployeeInfo = () => {
  const { id } = useParams();
  const navigate = useNavigate();
  const [employee, setEmployee] = useState(null);

  useEffect(() => {
    const fetchEmployee = async () => {
      try {
        const res = await fetch(`http://localhost:5000/api/employee/${id}`);
        const data = await res.json();
        if (!res.ok) throw new Error(data.message || 'Failed to fetch');
        setEmployee(data);
      } catch (err) {
        toast.error(err.message || 'Employee not found');
        navigate('/employee-info');
      }
    };

    fetchEmployee();
  }, [id, navigate]);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setEmployee((prev) => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();

    try {
      const response = await fetch(`http://localhost:5000/api/employee/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(employee),
      });

      if (!response.ok) throw new Error('Failed to update');
      toast.success('✅ Employee updated successfully!');
      setTimeout(() => navigate('/employee-info'), 1000);
    } catch (err) {
      toast.error(err.message || '❌ Update failed');
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
              type={dateFields.includes(key) ? 'date' : 'text'}
              id={key}
              name={key}
              value={
                dateFields.includes(key) && employee[key]
                  ? employee[key].slice(0, 10) // Format date
                  : employee[key] || ''
              }
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
