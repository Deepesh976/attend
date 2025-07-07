const Activity = require('../models/Activity');
const MonthlySummary = require('../models/MonthlySummary');
const XLSX = require('xlsx');

// Helper to normalize time to hh:mm:ss
const normalizeTime = (timeStr) => {
  if (!timeStr || typeof timeStr !== 'string') return '00:00:00';
  const parts = timeStr.split(':');
  if (parts.length === 2) return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}:00`;
  if (parts.length === 3) return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}:${parts[2].padStart(2, '0')}`;
  return '00:00:00';
};

// Helper to parse date in DD-MMM format
const parseExcelDate = (value) => {
  if (typeof value !== 'string' || !value.match(/^\d{1,2}-[A-Za-z]{3}$/)) return null;
  const [day, monthStr] = value.split('-');
  const monthMap = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
  };
  const month = monthMap[monthStr.toLowerCase()];
  if (!month) return null;
  const year = new Date().getFullYear(); // Assume current year; adjust if needed
  const fullDate = new Date(year, month - 1, parseInt(day));
  return isNaN(fullDate) ? null : fullDate;
};

// Helper to parse monthly summary data
const parseMonthlySummary = (summaryText, empId, empName, actualDates = []) => {
  console.log(`🔍 Raw summary text: "${summaryText}"`);
  
  // Clean the text but preserve the structure
  const cleanedText = summaryText.replace(/\s+/g, ' ').trim();
  console.log(`🧹 Cleaned text: "${cleanedText}"`);
  
  // More robust regex patterns to match the exact format from sample data
  const patterns = {
    // Patterns with " - " separator
    totalPresent: /Total Present\s*-\s*(\d+)/i,
    totalAbsent: /Total Absent\s*-\s*(\d+)/i,
    totalLeaveTaken: /Total Leave Taken\s*-\s*(\d+)/i,
    totalWeeklyOffPresent: /Total Weekly Off Present\s*-\s*(\d+)/i,
    totalDuration: /Total Duration\s*-\s*([\d:]+)/i,
    totalTDuration: /Total T\.?Duration\s*-\s*([\d:]+)/i,
    totalOverTime: /Total Over Time\s*-\s*([\d:]+)/i,
    totalLateBy: /Total LateBy\s*-\s*([\d:]+)(?:\s*\(Hrs\.\))?/i,
    totalEarlyBy: /Total EarlyBy\s*-\s*([\d:]+)(?:\s*\(Hrs\.\))?/i,
    totalRegularOT: /Total Regular OT\s*-\s*([\d:\-]+)(?:\s*\(Hrs\.\))?/i,
    
    // Patterns without " - " separator (space-separated)
    totalWOCount: /Total WO Count\s+(\d+)/i,
    totalHOCount: /Total HO Count\s+(\d+)/i
  };

  const extractedData = {};
  
  console.log(`🔍 Testing patterns against: "${cleanedText}"`);
  
  for (const [key, pattern] of Object.entries(patterns)) {
    const match = cleanedText.match(pattern);
    if (match) {
      console.log(`✅ ${key}: matched "${match[0]}" -> value: "${match[1]}"`);
      extractedData[key] = match[1];
    } else {
      console.log(`❌ ${key}: no match for pattern ${pattern}`);
    }
  }

  console.log(`📊 Extracted data:`, extractedData);

  // Helper function to extract date information from the actual date data (passed as parameter)
  const getDateInfo = (actualDates = []) => {
    const currentDate = new Date();
    let year = currentDate.getFullYear();
    let month = currentDate.getMonth() + 1;
    
    // If we have actual dates from the Excel data, use those to determine the primary month
    if (actualDates && actualDates.length > 0) {
      const monthCounts = {};
      let detectedYear = null;
      
      // Count occurrences of each month from actual dates
      actualDates.forEach(dateStr => {
        // Handle different date patterns from Excel
        let parsedDate = null;
        
        if (typeof dateStr === 'string') {
          // Pattern: DD-MMM (e.g., "21-Apr", "01-May")
          const monthMatch = dateStr.match(/(\d{1,2})-?([A-Za-z]{3})/);
          if (monthMatch) {
            const day = parseInt(monthMatch[1]);
            const monthAbbr = monthMatch[2].toLowerCase();
            
            const monthMap = {
              'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
              'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            };
            
            const monthNum = monthMap[monthAbbr];
            if (monthNum && day >= 1 && day <= 31) {
              parsedDate = new Date(currentDate.getFullYear(), monthNum - 1, day);
              if (!detectedYear) detectedYear = currentDate.getFullYear();
            }
          }
        } else if (typeof dateStr === 'number') {
          // Excel serial date
          parsedDate = new Date((dateStr - 25569) * 86400 * 1000);
          if (!detectedYear) detectedYear = parsedDate.getFullYear();
        }
        
        if (parsedDate && !isNaN(parsedDate.getTime())) {
          const monthKey = parsedDate.getMonth() + 1;
          monthCounts[monthKey] = (monthCounts[monthKey] || 0) + 1;
          console.log(`📅 Date "${dateStr}" -> Month ${monthKey}`);
        }
      });
      
      // Find the month with the most occurrences
      if (Object.keys(monthCounts).length > 0) {
        const primaryMonth = Object.keys(monthCounts).reduce((a, b) => 
          monthCounts[a] > monthCounts[b] ? a : b
        );
        
        month = parseInt(primaryMonth);
        if (detectedYear) year = detectedYear;
        
        console.log(`📅 Month counts:`, monthCounts);
        console.log(`📅 Primary month detected: ${month} (${year})`);
      }
    } else {
      // Fallback: try to extract month from the summary text
      const monthPatterns = [
        /(\d{1,2})-?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i,
        /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i
      ];
      
      for (const pattern of monthPatterns) {
        const match = summaryText.match(pattern);
        if (match) {
          const monthName = match[2] || match[1];
          const monthMap = {
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
          };
          const foundMonth = monthMap[monthName.toLowerCase()];
          if (foundMonth) {
            month = foundMonth;
            console.log(`📅 Detected month from summary text: ${monthName} (${month})`);
            break;
          }
        }
      }
    }
    
    return { year, month };
  };

  const { year, month } = getDateInfo(actualDates);
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // Clean up special values
  const cleanTimeValue = (value) => {
    if (!value) return '00:00';
    // Handle cases like "00:0-284" - extract the valid time part
    const timeMatch = value.match(/^(\d{1,2}:\d{2})/);
    return timeMatch ? timeMatch[1] : value;
  };

  const result = {
    empId,
    empName,
    year,
    month,
    monthName: monthNames[month - 1],
    totalPresent: parseInt(extractedData.totalPresent) || 0,
    totalAbsent: parseInt(extractedData.totalAbsent) || 0,
    totalLeaveTaken: parseInt(extractedData.totalLeaveTaken) || 0,
    totalWeeklyOffPresent: parseInt(extractedData.totalWeeklyOffPresent) || 0,
    totalDuration: cleanTimeValue(extractedData.totalDuration),
    totalTDuration: cleanTimeValue(extractedData.totalTDuration),
    totalOverTime: cleanTimeValue(extractedData.totalOverTime),
    totalWOCount: parseInt(extractedData.totalWOCount) || 0,
    totalHOCount: parseInt(extractedData.totalHOCount) || 0,
    totalLateBy: cleanTimeValue(extractedData.totalLateBy),
    totalEarlyBy: cleanTimeValue(extractedData.totalEarlyBy),
    totalRegularOT: cleanTimeValue(extractedData.totalRegularOT)
  };

  console.log(`💾 Final parsed summary:`, result);
  return result;
};

// Upload activity Excel
const uploadActivityExcel = async (req, res) => {
  try {
    if (!req.file || !req.file.buffer) {
      return res.status(400).json({ message: 'No file uploaded.' });
    }

    console.log('📊 Starting Excel processing...');
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    console.log(`📋 Total rows in Excel: ${rows.length}`);

    let activities = [];
    let monthlySummaries = [];
    let skippedRows = [];
    let i = 0;
    let employeeCount = 0;

    while (i < rows.length) {
      let empCode = '';
      let empName = '';
      let summary = {};
      let foundEmp = false;
      let startRow = i;

      console.log(`\n🔍 Searching for employee starting at row ${i}...`);

      // Find Employee Code and Name - More flexible search
      for (; i < rows.length && !foundEmp; i++) {
        const row = rows[i];
        for (let j = 0; j < row.length; j++) {
          if (typeof row[j] === 'string') {
            const cell = row[j].toLowerCase().replace(/\s/g, '');
            
            // More flexible employee code detection
            if ((cell.includes('employeecode') || cell.includes('empcode') || cell.includes('emp_code')) && !empCode) {
              // Look in same row first, then next cells
              for (let k = j; k < row.length; k++) {
                if (row[k] && row[k] !== row[j]) {
                  const potential = String(row[k]).trim();
                  if (potential && potential.length > 0 && potential !== 'undefined') {
                    empCode = potential;
                    console.log(`📧 Found employee code: "${empCode}" at row ${i}, col ${k}`);
                    break;
                  }
                }
              }
            }
            
            // More flexible employee name detection
            if ((cell.includes('employeename') || cell.includes('empname') || cell.includes('emp_name')) && !empName) {
              // Look in same row first, then next cells
              for (let k = j; k < row.length; k++) {
                if (row[k] && row[k] !== row[j]) {
                  const potential = String(row[k]).trim();
                  if (potential && potential.length > 2 && potential !== 'undefined') {
                    empName = potential;
                    console.log(`👤 Found employee name: "${empName}" at row ${i}, col ${k}`);
                    break;
                  }
                }
              }
            }
          }
        }
        if (empCode && empName) {
          foundEmp = true;
          employeeCount++;
          console.log(`✅ Employee ${employeeCount} found: ${empCode} - ${empName}`);
        }
      }

      if (!empCode || !empName) {
        console.log(`⚠️ No more employees found after row ${startRow}. Searched ${i - startRow} rows.`);
        break; // No more employees to process
      }

      // Find Dates Row first (Days row comes before summary in the actual format)
      let datesRowIdx = -1;
      for (; i < rows.length; i++) {
        if (rows[i][0]?.toString().trim().toLowerCase() === 'days') {
          datesRowIdx = i;
          console.log(`📅 Found 'Days' row at index ${i} for ${empCode}`);
          break;
        }
      }
      
      if (datesRowIdx === -1) {
        console.log(`❌ Days row not found for ${empCode} after row ${i}`);
        skippedRows.push({ row: i, reason: `Days row not found for ${empCode}` });
        continue;
      }

      // Now look for Summary Statistics AFTER the dates row but BEFORE the shift row
      let summaryFound = false;
      let monthlySummaryData = null;
      
      // Start searching from the row after dates row (should be day names, then summary)
      let summarySearchStart = datesRowIdx + 1;
      console.log(`🔍 Looking for summary data for ${empCode} starting from row ${summarySearchStart}`);
      
      for (let j = summarySearchStart; j < Math.min(summarySearchStart + 5, rows.length); j++) {
        const row = rows[j];
        console.log(`🔍 Checking row ${j} for summary data:`, row.slice(0, 3));
        
        // Check if any cell in the row contains "total present" (more flexible search)
        let summaryRowFound = false;
        let summaryText = '';
        
        for (let colIdx = 0; colIdx < Math.min(row.length, 5); colIdx++) {
          const cell = row[colIdx];
          if (cell && typeof cell === 'string') {
            const cellLower = cell.toLowerCase();
            if (cellLower.includes('total present')) {
              console.log(`📈 Found summary data at row ${j}, column ${colIdx} for ${empCode}`);
              console.log(`📝 Cell content: "${cell}"`);
              
              // If the entire summary is in one cell, use that
              if (cellLower.includes('total absent')) {
                summaryText = cell;
                summaryRowFound = true;
                console.log(`📊 Complete summary in single cell: "${summaryText}"`);
                break;
              } else {
                // Otherwise, combine the entire row
                summaryText = row.join(' ');
                summaryRowFound = true;
                console.log(`📊 Summary across row: "${summaryText}"`);
                break;
              }
            }
          }
        }
        
        if (summaryRowFound) {
          summaryFound = true;
          console.log(`✅ Found summary for ${empCode}!`);
          
          // Extract actual dates from the dates row to determine the correct month
          const datesRow = rows[datesRowIdx];
          const actualDates = [];
          
          console.log(`📅 Extracting dates from row ${datesRowIdx} for month detection...`);
          for (let col = 1; col < datesRow.length; col++) {
            const cellDate = datesRow[col];
            if (cellDate !== null && cellDate !== undefined && cellDate !== '') {
              actualDates.push(cellDate);
            }
          }
          
          console.log(`📅 Collected ${actualDates.length} dates:`, actualDates.slice(0, 5), '...');
          
          // Parse monthly summary from the text with actual dates
          monthlySummaryData = parseMonthlySummary(summaryText, empCode, empName, actualDates);
          
          // Add to monthlySummaries array
          if (monthlySummaryData) {
            monthlySummaries.push(monthlySummaryData);
            console.log(`✅ Successfully parsed and added monthly summary for ${empCode}`);
          }
          
          break;
        }
        
        // Stop looking for summary if we hit the next section (Shift row)
        if (rows[j][0]?.toString().trim().toLowerCase() === 'shift') {
          console.log(`📋 Found 'Shift' row at ${j}, stopping summary search`);
          break;
        }
      }

      if (!summaryFound) {
        console.log(`⚠️ No summary found for ${empCode} between rows ${summarySearchStart} and ${summarySearchStart + 5}`);
        console.log(`🔍 Debugging: Let's check what's in these rows:`);
        for (let debug = summarySearchStart; debug < Math.min(summarySearchStart + 5, rows.length); debug++) {
          const debugRow = rows[debug];
          console.log(`   Row ${debug}: [${debugRow.slice(0, 5).map(cell => `"${cell}"`).join(', ')}]`);
        }
      }

      
      const datesRow = rows[datesRowIdx];
      console.log(`📅 Found dates row for ${empCode} at index: ${datesRowIdx}`);

      // Find Data Rows (Shift, In Time, etc.)
      let headerRowIdx = -1;
      for (let j = datesRowIdx + 1; j < Math.min(datesRowIdx + 15, rows.length); j++) {
        if (rows[j][0]?.toString().trim().toLowerCase() === 'shift') {
          headerRowIdx = j;
          break;
        }
      }
      
      if (headerRowIdx === -1) {
        console.log(`❌ Shift row not found for ${empCode} after row ${datesRowIdx}`);
        skippedRows.push({ row: datesRowIdx, reason: `Shift row not found for ${empCode}` });
        // Try to find next employee
        i = datesRowIdx + 15; // Skip ahead to look for next employee
        continue;
      }

      const fields = rows.slice(headerRowIdx, headerRowIdx + 10);
      console.log(`📋 Field rows found for ${empCode}: ${fields.length}`);
      
      // Debug: Print all field labels found
      console.log(`🏷️ Field labels found for ${empCode}:`);
      fields.forEach((field, idx) => {
        const label = field[0]?.toString().trim().toLowerCase();
        if (label) {
          console.log(`  Row ${idx}: "${label}"`);
        }
      });

      const fieldRowMap = {};
      for (let idx = 0; idx < fields.length; idx++) {
        const fieldName = fields[idx][0]?.toString().trim().toLowerCase();
        if (fieldName) fieldRowMap[fieldName] = idx;
      }

      // More flexible field mapping
      const fieldKeys = [
        { key: 'shift', labels: ['shift'] },
        { key: 'timeInActual', labels: ['in time', 'time in', 'intime', 'in_time', 'timein'] },
        { key: 'timeOutActual', labels: ['out time', 'time out', 'outtime', 'out_time', 'timeout'] },
        { key: 'lateBy', labels: ['late by', 'late_by', 'lateby', 'late'] },
        { key: 'earlyBy', labels: ['early by', 'early_by', 'earlyby', 'early'] },
        { key: 'ot', labels: ['total ot', 'ot', 'overtime', 'over time'] },
        { key: 'duration', labels: ['duration', 'dur'] },
        { key: 'totalDuration', labels: ['t duration', 'total duration', 'total_duration', 'tduration'] },
        { key: 'status', labels: ['status'] }
      ];

      // Enhanced date column detection with comprehensive debugging
      const dateColIndices = [];
      console.log(`🔍 Analyzing date row for ${empCode}:`);
      console.log(`📊 Date row has ${datesRow.length} columns`);
      
      // Debug: show all cells in the date row
      for (let col = 0; col < Math.min(datesRow.length, 35); col++) {
        const cellValue = datesRow[col];
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          console.log(`  Col ${col}: "${cellValue}" (type: ${typeof cellValue})`);
        }
      }
      
      for (let col = 1; col < datesRow.length; col++) {
        const cellDate = datesRow[col];
        if (cellDate !== null && cellDate !== undefined && cellDate !== '') {
          const cellStr = String(cellDate).trim();
          
          // More comprehensive date pattern matching
          let isDateColumn = false;
          
          // Pattern 1: DD-MMM (e.g., 01-Jan, 15-Feb)
          if (cellStr.match(/^\d{1,2}-[A-Za-z]{3}$/)) {
            isDateColumn = true;
            console.log(`📅 Date pattern 1 match (DD-MMM): "${cellStr}"`);
          }
          // Pattern 2: DD/MM or MM/DD
          else if (cellStr.match(/^\d{1,2}\/\d{1,2}$/)) {
            isDateColumn = true;
            console.log(`📅 Date pattern 2 match (DD/MM): "${cellStr}"`);
          }
          // Pattern 3: DD-MM
          else if (cellStr.match(/^\d{1,2}-\d{1,2}$/)) {
            isDateColumn = true;
            console.log(`📅 Date pattern 3 match (DD-MM): "${cellStr}"`);
          }
          // Pattern 4: Just numbers (could be Excel date serial)
          else if (typeof cellDate === 'number' && cellDate > 40000 && cellDate < 50000) {
            isDateColumn = true;
            console.log(`📅 Date pattern 4 match (Excel serial): ${cellDate}`);
          }
          // Pattern 5: Single or double digit (likely day number)
          else if (cellStr.match(/^\d{1,2}$/) && parseInt(cellStr) >= 1 && parseInt(cellStr) <= 31) {
            isDateColumn = true;
            console.log(`📅 Date pattern 5 match (day number): "${cellStr}"`);
          }
          // Pattern 6: Any cell that looks like a date
          else if (cellStr.match(/\d/) && (cellStr.includes('-') || cellStr.includes('/') || cellStr.includes('.'))) {
            isDateColumn = true;
            console.log(`📅 Date pattern 6 match (general date): "${cellStr}"`);
          }
          
          if (isDateColumn) {
            dateColIndices.push(col);
          } else {
            console.log(`❌ Not a date: "${cellStr}" at column ${col}`);
          }
        }
      }
      
      console.log(`📊 Found ${dateColIndices.length} date columns for ${empCode}:`);
      console.log(`📊 Date column indices: [${dateColIndices.join(', ')}]`);

      if (dateColIndices.length === 0) {
        console.log(`⚠️ No date columns found for ${empCode}! This will result in 0 records.`);
        console.log(`🔍 Consider checking the Excel format - expected date row after 'Days' label`);
        i = headerRowIdx + 10;
        continue;
      }
      
      if (dateColIndices.length < 25) {
        console.log(`⚠️ Warning: Only ${dateColIndices.length} date columns found for ${empCode}. Expected around 30 for a full month.`);
      }

      const lateCountMap = {};
      const earlyGoMap = {};
      let recordsForThisEmployee = 0;

      for (const col of dateColIndices) {
        const cellDate = datesRow[col];
        const parsedDate = parseExcelDate(cellDate);
        if (!parsedDate) {
          // Try alternative date parsing
          const dateStr = String(cellDate).trim();
          console.log(`⚠️ Could not parse date: "${dateStr}" at column ${col}`);
          skippedRows.push({ row: datesRowIdx, col, reason: `Invalid date format: ${cellDate}` });
          continue;
        }

        const yearMonth = `${parsedDate.getFullYear()}-${parsedDate.getMonth() + 1}`;
        const empKey = `${empCode}|${yearMonth}`;
        if (!lateCountMap[empKey]) lateCountMap[empKey] = 0;
        if (!earlyGoMap[empKey]) earlyGoMap[empKey] = 0;

        const record = {
          date: parsedDate,
          empId: empCode,
          empName: empName,
          shift: '',
          timeInActual: '00:00:00',
          timeOutActual: '00:00:00',
          lateBy: '00:00:00',
          earlyBy: '00:00:00',
          ot: '00:00:00',
          duration: '00:00:00',
          totalDuration: '00:00:00',
          status: 'A',
          total_present: 0,
          total_absent: 0,
          total_leave: 0,
          total_wo: 0,
          total_ho: 0,
          total_regular_ot: '00:00:00'
        };

        // Enhanced field matching and value extraction
        for (const { key, labels } of fieldKeys) {
          let rowIdx = -1;
          let foundLabel = '';
          
          // Try to find matching label
          for (const label of labels) {
            if (fieldRowMap[label] !== undefined) {
              rowIdx = fieldRowMap[label];
              foundLabel = label;
              break;
            }
          }

          if (rowIdx !== -1) {
            let val = fields[rowIdx][col];
            
            // Handle different value types
            if (val === null || val === undefined) {
              val = '';
            } else if (typeof val === 'number') {
              // Handle Excel serial time values
              if (key === 'timeInActual' || key === 'timeOutActual') {
                if (val > 0 && val < 1) {
                  // Convert Excel serial time to HH:MM format
                  const totalMinutes = Math.round(val * 24 * 60);
                  const hours = Math.floor(totalMinutes / 60);
                  const minutes = totalMinutes % 60;
                  val = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
                } else {
                  val = val.toString();
                }
              } else {
                val = val.toString();
              }
            } else {
              val = val.toString();
            }
            
            val = val.trim();

            // Normalize time fields
            if (['timeInActual', 'timeOutActual', 'lateBy', 'earlyBy', 'ot', 'duration', 'totalDuration', 'total_regular_ot'].includes(key)) {
              val = normalizeTime(val);
            }

            if (key === 'status') {
              record.status = val || 'A';
            } else {
              record[key] = val;
            }
          }
        }

        // Weekend logic
        if (parsedDate.getDay() === 0) {
          record.shift = 'WO';
          record.status = 'WO';
          record.total_wo = 1;
        } else {
          const parseTime = t => {
            if (!t || t === '00:00:00') return null;
            const parts = t.split(':');
            if (parts.length < 2) return null;
            return Number(parts[0]) * 60 + Number(parts[1]);
          };
          
          const inMins = parseTime(record.timeInActual);
          const outMins = parseTime(record.timeOutActual);
          
          if (!inMins) {
            record.status = 'A';
            record.total_absent = 1;
          } else {
            if (inMins > 9 * 60 + 14 && inMins < 11 * 60) {
              lateCountMap[empKey]++;
              record.status = lateCountMap[empKey] <= 3 ? 'P' : '½P';
              record.total_present = record.status === 'P' ? 1 : 0.5;
            } else if (inMins >= 11 * 60) {
              record.status = '½P';
              record.total_present = 0.5;
            } else {
              record.status = 'P';
              record.total_present = 1;
            }

            if (outMins && outMins < 15 * 60 + 30) {
              earlyGoMap[empKey]++;
              if (earlyGoMap[empKey] > 2) {
                record.status = '½P';
                record.total_present = 0.5;
              }
            }
          }
        }

        activities.push(record);
        recordsForThisEmployee++;
      }

      console.log(`✅ Processed ${recordsForThisEmployee} records for ${empCode} - ${empName}`);

      // Move to next employee section - improved logic
      i = headerRowIdx + 10;
      
      console.log(`🔍 Looking for next employee starting from row ${i}...`);
      
      // Skip any remaining data rows for this employee
      let nextEmployeeFound = false;
      let searchEnd = Math.min(i + 50, rows.length); // Search next 50 rows maximum
      
      while (i < searchEnd) {
        const row = rows[i];
        if (row && row.length > 0 && row[0] && typeof row[0] === 'string') {
          const cell = row[0].toLowerCase().replace(/\s/g, '');
          if (cell.includes('employeecode') || cell.includes('empcode') || cell.includes('emp_code')) {
            console.log(`🎯 Found next employee marker at row ${i}: "${row[0]}"`);
            nextEmployeeFound = true;
            // Found next employee, don't increment i
            break;
          }
        }
        i++;
      }
      
      if (!nextEmployeeFound && i < rows.length) {
        console.log(`⚠️ No employee marker found in next 50 rows, continuing search from row ${i}`);
      } else if (i >= rows.length) {
        console.log(`📄 Reached end of file (row ${i}/${rows.length})`);
      }
    }

    console.log(`\n📊 COMPREHENSIVE PROCESSING SUMMARY:`);
    console.log(`=`.repeat(50));
    console.log(`📋 Excel Analysis:`);
    console.log(`  • Total rows in Excel: ${rows.length}`);
    console.log(`  • Rows processed: ${i}`);
    console.log(`  • Processing completion: ${((i/rows.length)*100).toFixed(1)}%`);
    
    console.log(`\n👥 Employee Analysis:`);
    console.log(`  • Employees found: ${employeeCount}`);
    console.log(`  • Expected employees: 5 (as per user)`);
    console.log(`  • Employee detection rate: ${((employeeCount/5)*100).toFixed(1)}%`);
    
    console.log(`\n📝 Record Analysis:`);
    console.log(`  • Total activity records created: ${activities.length}`);
    console.log(`  • Monthly summaries found: ${monthlySummaries.length}`);
    console.log(`  • Expected records: ~150 (5 employees × 30 days)`);
    console.log(`  • Record processing rate: ${((activities.length/150)*100).toFixed(1)}%`);
    console.log(`  • Average records per employee: ${employeeCount > 0 ? (activities.length/employeeCount).toFixed(1) : 0}`);
    
    console.log(`\n⚠️ Issues Found:`);
    console.log(`  • Skipped rows: ${skippedRows.length}`);
    if (skippedRows.length > 0) {
      console.log(`  • Skipped details:`, skippedRows.slice(0, 5)); // Show first 5 skipped rows
    }

    // Group by employee for detailed verification
    const byEmployee = activities.reduce((acc, record) => {
      if (!acc[record.empId]) acc[record.empId] = [];
      acc[record.empId].push(record);
      return acc;
    }, {});

    console.log(`\n📈 Detailed Records per Employee:`);
    if (Object.keys(byEmployee).length === 0) {
      console.log(`  ❌ No employees found in processed records!`);
    } else {
      Object.keys(byEmployee).forEach(empId => {
        const empRecords = byEmployee[empId].length;
        const status = empRecords >= 25 ? '✅' : empRecords >= 15 ? '⚠️' : '❌';
        console.log(`  ${status} ${empId}: ${empRecords} records`);
        
        // Show date range for this employee
        if (byEmployee[empId].length > 0) {
          const dates = byEmployee[empId].map(r => r.date).sort();
          const startDate = dates[0].toISOString().split('T')[0];
          const endDate = dates[dates.length-1].toISOString().split('T')[0];
          console.log(`    📅 Date range: ${startDate} to ${endDate}`);
        }
      });
    }

    console.log(`\n📊 Monthly Summaries Found:`);
    monthlySummaries.forEach(summary => {
      console.log(`  📈 ${summary.empId} (${summary.empName}): ${summary.monthName} ${summary.year}`);
      console.log(`     Present: ${summary.totalPresent}, Absent: ${summary.totalAbsent}`);
    });
    
    console.log(`\n🔍 Diagnosis:`);
    if (employeeCount < 5) {
      console.log(`  ❌ ISSUE: Missing employees (found ${employeeCount}/5)`);
      console.log(`     - Check Excel format for employee code/name markers`);
    }
    if (activities.length < 100) {
      console.log(`  ❌ ISSUE: Too few records (found ${activities.length}/150)`);
      console.log(`     - Check date column detection and parsing`);
    }
    if (employeeCount > 0 && activities.length > 0) {
      const avgRecordsPerEmp = activities.length / employeeCount;
      if (avgRecordsPerEmp < 25) {
        console.log(`  ❌ ISSUE: Low records per employee (avg: ${avgRecordsPerEmp.toFixed(1)})`);
        console.log(`     - Check date column detection patterns`);
      }
    }
    console.log(`=`.repeat(50));

    if (!activities.length) {
      return res.status(400).json({ 
        message: 'No valid activity data found in Excel.',
        skippedRows,
        employeeCount,
        monthlySummariesCount: monthlySummaries.length
      });
    }

    // Save activity records
    const bulkOps = activities.map(activity => ({
      updateOne: {
        filter: { empId: activity.empId, date: activity.date },
        update: { $set: activity },
        upsert: true
      }
    }));

    const activityResult = await Activity.bulkWrite(bulkOps);
    console.log(`✅ Saved ${activityResult.upsertedCount + activityResult.modifiedCount} activity records`);

    // Save monthly summaries
    let summaryResult = { upsertedCount: 0, modifiedCount: 0 };
    if (monthlySummaries.length > 0) {
      const summaryBulkOps = monthlySummaries.map(summary => ({
        updateOne: {
          filter: { 
            empId: summary.empId, 
            year: summary.year, 
            month: summary.month 
          },
          update: { $set: summary },
          upsert: true
        }
      }));

      summaryResult = await MonthlySummary.bulkWrite(summaryBulkOps);
      console.log(`✅ Saved ${summaryResult.upsertedCount + summaryResult.modifiedCount} monthly summaries`);
    }

    res.status(200).json({ 
      message: `Successfully processed ${employeeCount} employees with ${activities.length} activity records and ${monthlySummaries.length} monthly summaries.`,
      employeeCount,
      totalRecords: activities.length,
      monthlySummariesCount: monthlySummaries.length,
      activityInsertedCount: activityResult.upsertedCount + activityResult.modifiedCount,
      summaryInsertedCount: summaryResult.upsertedCount + summaryResult.modifiedCount,
      skippedRows,
      result: {
        activities: activityResult,
        summaries: summaryResult
      }
    });
  } catch (error) {
    console.error('❌ Excel Upload Error:', error);
    res.status(500).json({ 
      message: 'Server error while uploading activity Excel.', 
      error: error.message 
    });
  }
};

// Upload activity data (JSON format)
const uploadActivityData = async (req, res) => {
  try {
    const { activities } = req.body;

    if (!Array.isArray(activities) || activities.length === 0) {
      return res.status(400).json({ message: 'No activity data provided' });
    }

    const validActivities = [];
    const skippedRows = [];

    for (const activity of activities) {
      const empId = String(activity.empId || '').trim();
      const empName = String(activity.empName || '').trim();
      const date = activity.date;

      if (!empId || !empName || !date) {
        skippedRows.push(activity);
        continue;
      }

      const normalizedActivity = {
        ...activity,
        empId,
        empName,
        date: new Date(date),
        timeInActual: normalizeTime(activity.timeInActual),
        timeOutActual: normalizeTime(activity.timeOutActual),
        lateBy: normalizeTime(activity.lateBy),
        earlyBy: normalizeTime(activity.earlyBy),
        ot: normalizeTime(activity.ot),
        duration: normalizeTime(activity.duration),
        totalDuration: normalizeTime(activity.totalDuration),
        total_regular_ot: normalizeTime(activity.total_regular_ot),
        status: activity.status || 'A',
        total_present: activity.total_present || 0,
        total_absent: activity.total_absent || 0,
        total_leave: activity.total_leave || 0,
        total_wo: activity.total_wo || 0,
        total_ho: activity.total_ho || 0,
        summary_total_present: parseFloat(activity.summary_total_present || 0),
        summary_total_absent: parseFloat(activity.summary_total_absent || 0),
        summary_total_leave: parseFloat(activity.summary_total_leave || 0),
        summary_total_wo: parseFloat(activity.summary_total_wo || 0),
        summary_total_ho: parseFloat(activity.summary_total_ho || 0),
        summary_total_duration: activity.summary_total_duration || '00:00',
        summary_total_t_duration: activity.summary_total_t_duration || '00:00',
        summary_total_over_time: activity.summary_total_over_time || '00:00',
        summary_total_late_by: activity.summary_total_late_by || '00:00',
        summary_total_early_by: activity.summary_total_early_by || '00:00',
        summary_total_regular_ot: activity.summary_total_regular_ot || '00:00'
      };

      validActivities.push(normalizedActivity);
    }

    if (validActivities.length === 0) {
      return res.status(400).json({ 
        message: 'No valid activity data to insert', 
        skippedCount: skippedRows.length 
      });
    }

    const bulkOps = validActivities.map(activity => ({
      updateOne: {
        filter: { empId: activity.empId, date: activity.date },
        update: { $set: activity },
        upsert: true
      }
    }));

    const result = await Activity.bulkWrite(bulkOps);

    res.status(200).json({
      message: `Uploaded ${validActivities.length} activity records successfully`,
      insertedCount: result.upsertedCount + result.modifiedCount,
      skippedCount: skippedRows.length,
      result
    });
  } catch (error) {
    console.error('❌ Activity Upload Error:', error);
    res.status(500).json({ 
      message: 'Server error while uploading activity data', 
      error: error.message 
    });
  }
};

const getMonthlySummary = async (req, res) => {
  try {
    const activities = await Activity.find({});

    const grouped = {};

    for (const act of activities) {
      const {
        empId,
        empName,
        year,
        month,
        total_present = 0,
        total_absent = 0,
        total_leave = 0,
        total_wo = 0,
        total_ho = 0
      } = act;

      // Define April-May combo range
      const isAprilOrMay = month === 4 || month === 5;
      const monthGroupKey = isAprilOrMay ? 'Apr-May' : `M${month}`;
      const key = `${empId}_${year}_${monthGroupKey}`;

      if (!grouped[key]) {
        grouped[key] = {
          empId,
          empName,
          year,
          monthRange: monthGroupKey === 'Apr-May' ? 'April-May' : `Month ${month}`,
          totalPresent: 0,
          totalAbsent: 0,
          totalLeave: 0,
          totalWO: 0,
          totalHO: 0,
        };
      }

      grouped[key].totalPresent += total_present;
      grouped[key].totalAbsent += total_absent;
      grouped[key].totalLeave += total_leave;
      grouped[key].totalWO += total_wo;
      grouped[key].totalHO += total_ho;
    }

    const result = Object.values(grouped);

    res.status(200).json({
      success: true,
      count: result.length,
      data: result
    });
  } catch (error) {
    console.error('❌ getMonthlySummary error:', error);
    res.status(500).json({
      success: false,
      message: 'Server error while generating monthly summary',
      error: error.message
    });
  }
};

// Get all activities with enhanced filtering and pagination
const getAllActivities = async (req, res) => {
  try {
    const { 
      empId, 
      empName,
      status,
      startDate, 
      endDate, 
      page = 1, 
      limit, // No default limit - fetch all by default
      sortBy = 'date',
      sortOrder = 'desc'
    } = req.query;
    
    let filter = {};
    
    // Employee ID filter (exact match or partial)
    if (empId) {
      filter.empId = { $regex: empId, $options: 'i' };
    }
    
    // Employee name filter (partial match)
    if (empName) {
      filter.empName = { $regex: empName, $options: 'i' };
    }
    
    // Status filter (can be multiple statuses comma-separated)
    if (status) {
      const statuses = status.split(',').map(s => s.trim());
      filter.status = { $in: statuses };
    }
    
    // Date range filter
    if (startDate || endDate) {
      filter.date = {};
      if (startDate) {
        filter.date.$gte = new Date(startDate);
      }
      if (endDate) {
        filter.date.$lte = new Date(endDate);
      }
    }

    console.log(`🔍 Activity Query Filter:`, filter);
    
    // Build sort object
    const sortObj = {};
    sortObj[sortBy] = sortOrder === 'desc' ? -1 : 1;
    // Secondary sort by empId for consistency
    if (sortBy !== 'empId') {
      sortObj.empId = 1;
    }

    let query = Activity.find(filter).sort(sortObj);
    
    // Apply pagination only if limit is specified
    if (limit) {
      const skip = (parseInt(page) - 1) * parseInt(limit);
      query = query.skip(skip).limit(parseInt(limit));
    }
    
    const activities = await query;
    const totalCount = await Activity.countDocuments(filter);
    
    console.log(`📊 Query Results: ${activities.length} activities found (Total: ${totalCount})`);
    
    // Response format
    const response = {
      activities,
      totalCount,
      currentPage: parseInt(page),
      hasMore: false
    };
    
    // Add pagination info only if limit was used
    if (limit) {
      const limitNum = parseInt(limit);
      response.totalPages = Math.ceil(totalCount / limitNum);
      response.hasMore = (parseInt(page) * limitNum) < totalCount;
      response.limit = limitNum;
    } else {
      response.totalPages = 1;
      response.limit = totalCount; // All records fetched
    }
    
    res.status(200).json(response);
  } catch (error) {
    console.error('❌ Get Activities Error:', error);
    res.status(500).json({ 
      message: 'Server error while fetching activities', 
      error: error.message 
    });
  }
};

// Delete all activities
const deleteAllActivities = async (req, res) => {
  try {
    const result = await Activity.deleteMany({}); // 🔥 Deletes all records
    console.log(`🗑️ Deleted ${result.deletedCount} activity records`);

    res.status(200).json({
      message: `${result.deletedCount} activity records deleted successfully`,
      deletedCount: result.deletedCount
    });
  } catch (error) {
    console.error('❌ Delete Activities Error:', error);
    res.status(500).json({
      message: 'Server error while deleting all activities',
      error: error.message
    });
  }
};

module.exports = {
  uploadActivityData,
  getAllActivities,
  deleteAllActivities,
  uploadActivityExcel,
  getMonthlySummary
};