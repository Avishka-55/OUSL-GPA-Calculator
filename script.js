document.getElementById('fileInput').addEventListener('change', function (e) {
  const fileInput = e.target;
  const label = document.getElementById('fileLabelText');
  if (fileInput.files && fileInput.files.length > 0) {
    label.textContent = fileInput.files[0].name;
  } else {
    label.textContent = 'No file chosen';
  }
});

document.getElementById('calculateBtn').addEventListener('click', function () {
  const fileInput = document.getElementById('fileInput');
  const file = fileInput.files[0];
  const resultDiv = document.getElementById('result');
  const nongpaInput = document.getElementById('nongpaInput').value;
  const nonGpaCourses = nongpaInput.split(',').map(code => code.trim().toUpperCase()).filter(code => code);

  if (!file) {
    resultDiv.innerHTML = '<span style="color:#c0392b;">Please upload an .xlsx or .xls file first.</span>';
    return;
  }

  resultDiv.innerHTML = 'Calculating...';

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const courses = XLSX.utils.sheet_to_json(sheet);

      const gpa = calculateGPA(courses, nonGpaCourses);
      displayResult(gpa);
    } catch (err) {
      resultDiv.innerHTML = '<span style="color:#c0392b;">Error processing file. Please make sure it is a valid Excel result sheet.</span>';
      console.error(err);
    }
  };
  reader.readAsBinaryString(file);
});

function calculateGPA(courses, nonGpaCourses) {
  let totalCredits = 0;
  let weightedGradePoints = 0;

  courses.forEach(course => {
    const code = course['Course Code'] ? course['Course Code'].toUpperCase() : '';
    const grade = course['Grade'];

     if (!grade || grade === '-' || grade.toLowerCase() === 'pending' || grade.toLowerCase() === 'rx' || grade.toLowerCase() === 'fa' ||
        grade.toLowerCase() === 'n/a' || grade.toLowerCase() === 'p') {
      return;
    }


    if (nonGpaCourses.includes(code)) {
      return;
    }

    const gradePoint = getGradePoint(grade);
    const credit = getCreditFromCode(code);

    if (credit && gradePoint !== undefined) {
      totalCredits += credit;
      weightedGradePoints += gradePoint * credit;
    }
  });

  if (totalCredits === 0) {
    return { error: 'Error: No GPA courses available or data is invalid' };
  }

  const gpa = (weightedGradePoints / totalCredits).toFixed(3);
  return { gpa, totalCredits };
}

function getGradePoint(grade) {
  const gradeMapping = {
    'A+': 4.0,
    'A': 4.0,
    'A-': 3.7,
    'B+': 3.3,
    'B': 3.0,
    'B-': 2.7,
    'C+': 2.3,
    'C': 2.0,
    'C-': 1.7,
    'D+': 1.3,
    'D': 1.0,
    'E': 0.0,
    'F': 0.0
  };
  return gradeMapping[grade] !== undefined ? gradeMapping[grade] : 0.0;
}

function getCreditFromCode(courseCode) {
  const creditChar = courseCode ? courseCode.charAt(4) : '';
  const creditValue = parseInt(creditChar, 10);
  return creditValue;
}

function getClassFromGPA(gpa) {
  if (gpa >= 3.7) {
    return "First Class";
  } else if (gpa >= 3.3) {
    return "Second Class Upper";
  } else if (gpa >= 3.0) {
    return "Second Class Lower";
  } else if (gpa >= 2.0) {
    return "General Pass";
  } else if (gpa >= 0) {
    return "Below Pass";
  } else {
    return "Invalid GPA";
  }
}

function displayResult(result) {
  const resultDiv = document.getElementById('result');
  if (result.error) {
    resultDiv.innerHTML = `<span style="color:#c0392b;">${result.error}</span>`;
  } else {
    const gpaValue = parseFloat(result.gpa);
    const className = getClassFromGPA(gpaValue);
    resultDiv.innerHTML = `<span style="
  display: inline-block;
  background-color: #fff9f0;
  border-left: 6px solid #FFB347;
  padding: 16px 24px;
  font-family: 'Arial Rounded MT', sans-serif;
  font-size: 16px;
  color: #444;
  text-align: left;
  border-radius: 10px;
  box-shadow: 0 3px 10px rgba(0,0,0,0.05);
">
  🏆 <b style="color:blue;">Your GPA Breakdown</b><br><br>
  ✅ Total Credits Counted: <b style="color:#F4C430;">${result.totalCredits}</b><br>
  📊 Final GPA Score: <b style="color:#F4C430;">${result.gpa}</b><br>
  🏅 Awarded Class: <b style="color:#F4C430;">${className}</b>
</span>

`;
  }
}
