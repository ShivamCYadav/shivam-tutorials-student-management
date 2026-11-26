
async function exportBackup() {
  await DB.init();

  const students = await DB.getAll("students");
  const payments = await DB.getAll("payments");
  const attendance = await DB.getAll("attendance");

  const workbook = XLSX.utils.book_new();

  const studentRows = students.map(s => ({
    "Student ID": s.id,
    "Student Code": s.studentCode,
    "Name": s.name,
    "Class": s.className,
    "Board": s.board,
    "WhatsApp": s.whatsapp,
    "Total Fees": s.totalFees,
    "Admission Date": s.admissionDate,
    "Status": s.status
  }));

  const feeRows = payments.map(p => ({
    "Payment ID": p.id,
    "Student ID": p.studentId,
    "Student Name": p.studentName,
    "Date": p.date,
    "Amount": p.amount,
    "Mode": p.mode,
    "Notes": p.notes
  }));

  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(studentRows),
    "Students"
  );

  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(feeRows),
    "Fees"
  );

  // Group attendance records by month.
  const monthGroups = {};

  attendance.forEach(record => {
    const monthKey = record.date.slice(0, 7);

    if (!monthGroups[monthKey]) {
      monthGroups[monthKey] = [];
    }

    monthGroups[monthKey].push(record);
  });

  Object.entries(monthGroups).forEach(([monthKey, records]) => {
    const [year, month] = monthKey.split("-");
    const daysInMonth = new Date(Number(year), Number(month), 0).getDate();

    const monthStudents = students.filter(s =>
      records.some(r => r.studentId === s.id)
    );

    const rows = monthStudents.map(student => {
      const row = {
        "Student ID": student.id,
        "Student Code": student.studentCode,
        "Student Name": student.name,
        "Class": student.className
      };

      for (let day = 1; day <= daysInMonth; day++) {
        const date = monthKey + "-" + String(day).padStart(2, "0");

        const record = records.find(r =>
          r.studentId === student.id && r.date === date
        );

        row[String(day).padStart(2, "0") + " " + monthKey] =
          record ? (record.status === "Present" ? "P" : "A") : "";
      }

      return row;
    });

    const sheetName = "Att_" + year + "_" + month;
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(rows),
      sheetName
    );
  });

  // Add an attendance sheet even if no attendance has been recorded.
  if (Object.keys(monthGroups).length === 0) {
    const current = new Date();
    const monthKey = current.getFullYear() + "-" +
      String(current.getMonth() + 1).padStart(2, "0");

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["Student ID", "Student Code", "Student Name", "Class"]
      ]),
      "Att_" + monthKey.replace("-", "_")
    );
  }

  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, "Shivam_Tutorials_Backup_" + date + ".xlsx");
}

async function importBackup(file) {
  const result = await Swal.fire({
    title: "Restore this backup?",
    text: "Current student, fee and attendance records will be replaced.",
    icon: "warning",
    showCancelButton: true,
    confirmButtonText: "Yes, restore backup"
  });

  if (!result.isConfirmed) return;

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const requiredSheets = ["Students", "Fees"];

  for (const sheet of requiredSheets) {
    if (!workbook.Sheets[sheet]) {
      throw new Error("This workbook is missing the " + sheet + " sheet.");
    }
  }

  const students = XLSX.utils.sheet_to_json(workbook.Sheets.Students);
  const fees = XLSX.utils.sheet_to_json(workbook.Sheets.Fees);

  const attendance = [];

  workbook.SheetNames.filter(name => name.startsWith("Att_"))
    .forEach(sheetName => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        defval: ""
      });

      rows.forEach(row => {
        Object.entries(row).forEach(([column, value]) => {
          const match = column.match(/^(\d{2}) (\d{4}-\d{2})$/);
          if (!match || !value) return;

          const date = match[2] + "-" + match[1];
          const status = value === "P" ? "Present" :
                         value === "A" ? "Absent" : null;

          if (!status) return;

          attendance.push({
            id: date + "_" + row["Student ID"],
            studentId: row["Student ID"],
            studentCode: row["Student Code"],
            studentName: row["Student Name"],
            className: row["Class"],
            date,
            status
          });
        });
      });
    });

  const restoredStudents = students.map(s => ({
    id: String(s["Student ID"]),
    studentCode: String(s["Student Code"] || ""),
    name: String(s["Name"] || ""),
    className: String(s["Class"] || ""),
    board: String(s["Board"] || ""),
    whatsapp: String(s["WhatsApp"] || ""),
    totalFees: Number(s["Total Fees"] || 0),
    admissionDate: String(s["Admission Date"] || ""),
    status: String(s["Status"] || "Active")
  }));

  const restoredFees = fees.map(p => ({
    id: String(p["Payment ID"]),
    studentId: String(p["Student ID"]),
    studentName: String(p["Student Name"] || ""),
    date: String(p["Date"] || ""),
    amount: Number(p["Amount"] || 0),
    mode: String(p["Mode"] || ""),
    notes: String(p["Notes"] || "")
  }));

  const validStudentIds = new Set(restoredStudents.map(s => s.id));

  if (restoredFees.some(p => !validStudentIds.has(p.studentId)) ||
      attendance.some(a => !validStudentIds.has(a.studentId))) {
    throw new Error("The backup contains records linked to unknown students.");
  }

  // Validate all data before clearing the current database.
  if (restoredStudents.some(s => !s.id || !s.name)) {
    throw new Error("The backup contains an invalid student record.");
  }

  await DB.clear("students");
  await DB.clear("payments");
  await DB.clear("attendance");

  for (const student of restoredStudents) {
    await DB.put("students", student);
  }

  for (const payment of restoredFees) {
    await DB.put("payments", payment);
  }

  for (const record of attendance) {
    await DB.put("attendance", record);
  }

  await Swal.fire(
    "Backup Restored",
    "Student, fee and attendance records have been imported.",
    "success"
  );
}

document.getElementById("exportButton")
  .addEventListener("click", () => {
    exportBackup().catch(error =>
      Swal.fire("Export Failed", error.message, "error")
    );
  });

document.getElementById("importButton")
  .addEventListener("click", () => {
    const file = document.getElementById("importFile").files[0];

    if (!file) {
      Swal.fire("Select a file", "Please choose an Excel backup.", "warning");
      return;
    }

    importBackup(file).catch(error =>
      Swal.fire("Import Failed", error.message, "error")
    );
  });
