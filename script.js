const PASS_MARK = 9;
const A_PLUS_MARK = 27;

let reportTitle = "Result Summary";

/* =====================================================
EVENTS
===================================================== */

document.getElementById("fileInput").addEventListener("change", handleFile);

document.getElementById("downloadBtn").addEventListener("click", downloadPDF);

/* =====================================================
FILE UPLOAD
===================================================== */

async function handleFile(e) {
  const data = await e.target.files[0].arrayBuffer();

  const workbook = XLSX.read(data);

  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  /* ---------- TITLE ---------- */

  const excelTitle = raw[0][0] || "Result Summary";

  reportTitle = excelTitle.replace(/Mark\s*List/i, "Summary");

  /* ---------- STUDENTS ---------- */

  const rows = raw.slice(2).filter((r) => r[1]);

  const students = rows.map((r) => ({
    name: r[1],

    english: parse(r[2]),
    malayalam: parse(r[3]),
    hindi: parse(r[4]),

    business: parse(r[5]),
    accountancy: parse(r[6]),
    economics: parse(r[7]),
    ca: parse(r[8]),
  }));

  render(students);
}

/* =====================================================
HELPERS
===================================================== */

function parse(v) {
  if (v === "ab" || v === "AB" || !v) {
    return {
      mark: 0,
      isAbsent: true,
    };
  }

  return {
    mark: Number(v) || 0,
    isAbsent: false,
  };
}

function language(student) {
  return !student.malayalam.isAbsent ? student.malayalam : student.hindi;
}

function subjects(student) {
  return {
    English: student.english,
    Language: language(student),

    Business: student.business,
    Accountancy: student.accountancy,
    Economics: student.economics,
    CA: student.ca,
  };
}

function total(student) {
  return Object.values(subjects(student)).reduce((sum, s) => sum + s.mark, 0);
}

function status(student) {
  const subs = Object.values(subjects(student));

  if (subs.some((s) => s.isAbsent)) {
    return "ABSENT";
  }

  if (subs.every((s) => s.mark >= PASS_MARK)) {
    return "PASS";
  }

  return "FAIL";
}

/* =====================================================
GRADE SYSTEM
===================================================== */

function subjectGrade(mark) {
  if (mark >= 27) return "A+";
  if (mark >= 24) return "A";
  if (mark >= 21) return "B+";
  if (mark >= 18) return "B";
  if (mark >= 15) return "C+";
  if (mark >= 12) return "C";
  if (mark >= 9) return "D+";

  return "D";
}

/* =====================================================
MAIN RENDER
===================================================== */

function render(students) {
  document.getElementById("dynamicTitle").innerText = reportTitle;

  renderExecutiveSummary(students);

  renderFailure(students);

  renderAbsence(students);

  renderAPlus(students);

  renderNearPass(students);

  renderRisk(students);

  renderRank(students);

  renderAbsent(students);
}

/* =====================================================
EXECUTIVE SUMMARY
===================================================== */

function renderExecutiveSummary(students) {
  const totalStudents = students.length;

  const eligible = students.filter((s) => status(s) === "PASS").length;

  const notEligible = totalStudents - eligible;

  const percentage = ((eligible / totalStudents) * 100).toFixed(2);

  const fullAPlus = students.filter((student) => {
    return Object.values(subjects(student)).every(
      (s) => subjectGrade(s.mark) === "A+"
    );
  }).length;

  document.getElementById("executiveSummary").innerHTML = `

  <table>

    <tr>
      <th>Total Students Registered</th>
      <td>${totalStudents}</td>
    </tr>

    <tr>
      <th>Eligible For Higher Studies</th>
      <td>${eligible}</td>
    </tr>

    <tr>
      <th>Not Eligible</th>
      <td>${notEligible}</td>
    </tr>

    <tr>
      <th>Pass Percentage</th>
      <td>${percentage}%</td>
    </tr>

    <tr>
      <th>No. of Full A+ Students</th>
      <td>${fullAPlus}</td>
    </tr>

  </table>

  `;

  renderSubjectGradeMatrix(students);

  renderAPlusSummary(students);
}

/* =====================================================
SUBJECT GRADE MATRIX
===================================================== */

function renderSubjectGradeMatrix(students) {
  const grades = ["A+", "A", "B+", "B", "C+", "C", "D+", "D"];

  const subs = Object.keys(subjects(students[0]));

  let html = `

  <h3>📊 Subject-wise Grade Distribution</h3>

  <table>

    <tr>
      <th>Grade</th>

      ${subs.map((s) => `<th>${s}</th>`).join("")}

    </tr>

  `;

  grades.forEach((grade) => {
    html += `<tr><td>${grade}</td>`;

    subs.forEach((sub) => {
      const count = students.filter((student) => {
        const mark = subjects(student)[sub].mark;

        return subjectGrade(mark) === grade;
      }).length;

      html += `<td>${count}</td>`;
    });

    html += "</tr>";
  });

  html += "</table>";

  document.getElementById("subjectGradeMatrix").innerHTML = html;
}

/* =====================================================
A+ SUMMARY
===================================================== */

function renderAPlusSummary(students) {
  const counts = {
    6: 0,
    5: 0,
    4: 0,
    3: 0,
    2: 0,
    1: 0,
  };

  students.forEach((student) => {
    let aplus = 0;

    Object.values(subjects(student)).forEach((sub) => {
      if (subjectGrade(sub.mark) === "A+") {
        aplus++;
      }
    });

    if (aplus >= 1) {
      counts[aplus]++;
    }
  });

  let html = `

  <h3>⭐ A+ Distribution</h3>

  <table>

    <tr>
      <th>Category</th>
      <th>Count</th>
    </tr>

  `;

  Object.entries(counts)
    .sort((a, b) => b[0] - a[0])
    .forEach(([k, v]) => {
      html += `

      <tr>
        <td>Students with ${k} A+</td>
        <td>${v}</td>
      </tr>

      `;
    });

  html += "</table>";

  document.getElementById("aplusSummary").innerHTML = html;
}

/* =====================================================
SUBJECT PASS %
===================================================== */

function renderFailure(students) {
  let html = `

  <h3>📈 Subject Pass %</h3>

  <table>

    <tr>
      <th>Subject</th>
      <th>Pass %</th>
    </tr>

  `;

  Object.keys(subjects(students[0])).forEach((sub) => {
    const appeared = students.filter((s) => !subjects(s)[sub].isAbsent).length;

    const passed = students.filter(
      (s) => !subjects(s)[sub].isAbsent && subjects(s)[sub].mark >= PASS_MARK
    ).length;

    const percent = appeared ? ((passed / appeared) * 100).toFixed(1) : 0;

    html += `

    <tr>
      <td>${sub}</td>
      <td>${percent}%</td>
    </tr>

    `;
  });

  html += "</table>";

  document.getElementById("subjectFailure").innerHTML = html;
}

/* =====================================================
SUBJECT ABSENCE
===================================================== */

function renderAbsence(students) {
  let html = `

  <h3>🚫 Subject-wise Absence</h3>

  <table>

    <tr>
      <th>Subject</th>
      <th>Absent Count</th>
    </tr>

  `;

  Object.keys(subjects(students[0])).forEach((sub) => {
    const count = students.filter((s) => subjects(s)[sub].isAbsent).length;

    html += `

    <tr>
      <td>${sub}</td>
      <td>${count}</td>
    </tr>

    `;
  });

  html += "</table>";

  document.getElementById("subjectAbsence").innerHTML = html;
}

/* =====================================================
A+ ANALYSIS
===================================================== */

function renderAPlus(students) {
  let fullAPlus = [];
  let nearAPlus = [];

  students.forEach((student) => {
    const weak = [];

    Object.entries(subjects(student)).forEach(([sub, val]) => {
      if (val.mark < A_PLUS_MARK) {
        weak.push({
          subject: sub,
          missing: A_PLUS_MARK - val.mark,
        });
      }
    });

    if (weak.length === 0) {
      fullAPlus.push(student.name);
    } else if (weak.length <= 2 && status(student) !== "ABSENT") {
      nearAPlus.push({
        name: student.name,
        weak,
      });
    }
  });

  let html = `

  <h3>⭐ A+ Analysis</h3>

  <b>Full A+ Students:</b><br><br>

  ${fullAPlus.length ? fullAPlus.join("<br>") : "None"}

  <br><br>

  <table>

    <tr>
      <th>Name</th>
      <th>Near A+ Subjects</th>
    </tr>

  `;

  nearAPlus.forEach((student) => {
    const details = student.weak
      .map(
        (w) =>
          `${w.subject}
        (${w.missing} marks needed)`
      )
      .join(", ");

    html += `

    <tr>
      <td>${student.name}</td>
      <td>${details}</td>
    </tr>

    `;
  });

  html += "</table>";

  document.getElementById("aplusPanel").innerHTML = html;
}

/* =====================================================
NEAR PASS
===================================================== */

function renderNearPass(students) {
  let list = [];

  students.forEach((student) => {
    Object.entries(subjects(student)).forEach(([sub, val]) => {
      if (!val.isAbsent && val.mark >= PASS_MARK - 2 && val.mark < PASS_MARK) {
        list.push({
          name: student.name,
          subject: sub,
          mark: val.mark,
        });
      }
    });
  });

  let html = `

  <h3>📉 Near Pass / Borderline Students</h3>

  <table>

    <tr>
      <th>Name</th>
      <th>Subject</th>
      <th>Mark</th>
    </tr>

  `;

  list.forEach((student) => {
    html += `

    <tr>
      <td>${student.name}</td>
      <td>${student.subject}</td>
      <td>${student.mark}</td>
    </tr>

    `;
  });

  html += "</table>";

  document.getElementById("nearPassPanel").innerHTML = html;
}

/* =====================================================
AT RISK
===================================================== */

function renderRisk(students) {
  let risk = [];

  students.forEach((student) => {
    const failedSubjects = Object.entries(subjects(student))
      .filter(([k, v]) => !v.isAbsent && v.mark < PASS_MARK)
      .map(([k]) => k);

    let level = "";
    let reason = "";

    if (failedSubjects.length >= 3) {
      level = "HIGH";

      reason = `Failed ${failedSubjects.length} subjects`;
    } else if (failedSubjects.length === 2) {
      level = "MEDIUM";

      reason = "Failed 2 subjects";
    } else if (total(student) < 60) {
      level = "MEDIUM";

      reason = "Very low total score";
    }

    if (level) {
      risk.push({
        name: student.name,
        level,
        reason,
      });
    }
  });

  let html = `

  <h3>🚨 At-Risk Students</h3>

  <table>

    <tr>
      <th>Name</th>
      <th>Risk Level</th>
      <th>Reason</th>
    </tr>

  `;

  risk.forEach((student) => {
    html += `

    <tr>
      <td>${student.name}</td>
      <td>${student.level}</td>
      <td>${student.reason}</td>
    </tr>

    `;
  });

  html += "</table>";

  document.getElementById("riskPanel").innerHTML = html;
}

/* =====================================================
RANK LIST
===================================================== */

function renderRank(students) {
  const ranked = students
    .filter((s) => status(s) !== "ABSENT")
    .map((s) => ({
      name: s.name,
      total: total(s),
    }))
    .sort((a, b) => b.total - a.total);

  let html = `

  <h3>🏆 Rank List</h3>

  <b>Top 5</b><br><br>

  `;

  ranked.slice(0, 5).forEach((s, i) => {
    html += `
      ${i + 1}. ${s.name} (${s.total})<br>
      `;
  });

  html += `
  <br><br>
  <b>Bottom 5</b><br><br>
  `;

  ranked.slice(-5).forEach((s) => {
    html += `
      ${s.name} (${s.total})<br>
      `;
  });

  document.getElementById("rankList").innerHTML = html;
}

/* =====================================================
ABSENTEE REPORT
===================================================== */

function renderAbsent(students) {
  const absentees = students
    .map((student) => {
      const missed = Object.entries(subjects(student))
        .filter(([k, v]) => v.isAbsent)
        .map(([k]) => k);

      return {
        name: student.name,
        missed,
      };
    })
    .filter((x) => x.missed.length > 0);

  let html = `

  <h3>🚫 Absentee Report</h3>

  <table>

    <tr>
      <th>Name</th>
      <th>Absent Subjects</th>
    </tr>

  `;

  absentees.forEach((student) => {
    html += `

    <tr>

      <td>${student.name}</td>

      <td>

      ${
        student.missed.length === 6 ? "All Subjects" : student.missed.join(", ")
      }

      </td>

    </tr>

    `;
  });

  html += "</table>";

  document.getElementById("absentPanel").innerHTML = html;
}

/* =====================================================
PDF DOWNLOAD
===================================================== */

async function downloadPDF() {
  const element = document.getElementById("report");

  const options = {
    margin: [8, 8, 8, 8],

    filename: "Result_Analysis.pdf",

    image: {
      type: "jpeg",
      quality: 1,
    },

    html2canvas: {
      scale: 2,
      useCORS: true,
    },

    jsPDF: {
      unit: "mm",
      format: "a4",
      orientation: "portrait",
    },

    pagebreak: {
      mode: ["css"],
    },
  };

  await html2pdf().set(options).from(element).save();
}
