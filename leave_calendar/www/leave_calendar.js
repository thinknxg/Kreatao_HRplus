function today() {
  const currentDay = document.getElementById("today");
  if (currentDay) {
    currentDay.scrollIntoView({ inline: "center", block: "nearest" });
  }
}
today();

function changeYear(direction) {
  var currentYear = parseInt(document.querySelector(".current-year-display").textContent);
  var newYear = currentYear + direction;
  var url = new URL(window.location.href);
  url.searchParams.set("year", newYear);
  url.searchParams.set("from_date", newYear + "-01-01");
  url.searchParams.set("to_date", newYear + "-12-31");
  window.location.href = url.toString();
}

function updateFilters() {
  var department = document.getElementById("departmentFilter").value;
  var employee = document.getElementById("employeeFilter").value;
  var fromDate = document.getElementById("fromDate").value;
  var toDate = document.getElementById("toDate").value;
  var url = new URL(window.location.href);
  if (department !== "all") { url.searchParams.set("department", department); } else { url.searchParams.delete("department"); }
  if (employee !== "all") { url.searchParams.set("employee", employee); } else { url.searchParams.delete("employee"); }
  if (fromDate) { url.searchParams.set("from_date", fromDate); }
  if (toDate) { url.searchParams.set("to_date", toDate); }
  window.location.href = url.toString();
}

function getCalendarData() {
  var el = document.getElementById("calendarData");
  return JSON.parse(el.textContent);
}

function hexToRgb(hex) {
  hex = (hex || "#ffffff").replace("#", "");
  if (hex.length === 3) hex = hex.split("").map(function(c) { return c + c; }).join("");
  var r = parseInt(hex.substring(0, 2), 16) || 255;
  var g = parseInt(hex.substring(2, 4), 16) || 255;
  var b = parseInt(hex.substring(4, 6), 16) || 255;
  return [r, g, b];
}

function exportToExcel() {
  if (window.XLSX) { doExcelExport(); return; }
  var script = document.createElement("script");
  script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  script.onload = doExcelExport;
  document.head.appendChild(script);
}

function doExcelExport() {
  var data = getCalendarData();
  var wb = XLSX.utils.book_new();
  var wsData = [];

  wsData.push(data.headers);

  data.rows.forEach(function(row) {
    var r = [row.name, row.total, row.used, row.remaining, row.new_leaves, row.carry];
    row.days.forEach(function(day) { r.push(day.abbr || ""); });
    wsData.push(r);
  });

  var ws = XLSX.utils.aoa_to_sheet(wsData);

  ws["!cols"] = data.headers.map(function(h, i) {
    return { wch: i === 0 ? 25 : i < 6 ? 8 : 5 };
  });

  XLSX.utils.book_append_sheet(wb, ws, "Leave Calendar");
  var fromDate = document.getElementById("fromDate").value;
  var toDate = document.getElementById("toDate").value;
  XLSX.writeFile(wb, "leave_calendar_" + fromDate + "_to_" + toDate + ".xlsx");
}

function exportToPDF() {
  if (window.jspdf) { doLoadAutoTable(); return; }
  var script1 = document.createElement("script");
  script1.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
  script1.onload = doLoadAutoTable;
  document.head.appendChild(script1);
}

function doLoadAutoTable() {
  if (window.jspdf && window.jspdf.jsPDF.prototype.autoTable) { doPDFExport(); return; }
  var script2 = document.createElement("script");
  script2.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js";
  script2.onload = doPDFExport;
  document.head.appendChild(script2);
}

function doPDFExport() {
  var data = getCalendarData();
  var fromDate = document.getElementById("fromDate").value;
  var toDate = document.getElementById("toDate").value;
  var { jsPDF } = window.jspdf;

  var doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
  doc.setFontSize(13);
  doc.setTextColor(40, 40, 40);
  doc.text("Leave Calendar: " + fromDate + " to " + toDate, 14, 12);

  var head = [data.headers];
  var body = data.rows.map(function(row) {
    var r = [
      row.name,
      String(row.total !== undefined ? row.total : ""),
      String(row.used !== undefined ? row.used : ""),
      String(row.remaining !== undefined ? row.remaining : ""),
      String(row.new_leaves !== undefined ? row.new_leaves : ""),
      String(row.carry !== undefined ? row.carry : "")
    ];
    row.days.forEach(function(day) { r.push(day.abbr || ""); });
    return r;
  });

  // Build per-cell color map
  var cellColorMap = {};
  data.rows.forEach(function(row, rowIdx) {
    row.days.forEach(function(day, colIdx) {
      var col = colIdx + 6;
      var rgb;
      if (!day.name || day.name === "DEFAULT") {
        rgb = [255, 255, 255];
      } else if (day.name === "SATURDAY" || day.name === "SUNDAY") {
        rgb = [189, 189, 189];
      } else if (day.name === "OFFICIAL HOLIDAY") {
        rgb = [120, 202, 121];
      } else if (day.name === "ABSENCE") {
        rgb = [202, 63, 63];
      } else {
        rgb = hexToRgb(day.color);
      }
      cellColorMap[rowIdx + "_" + col] = rgb;
    });
  });

  // Dynamic column styles — fixed width for first 6, auto for day columns
  var colStyles = {
    0: { cellWidth: 30, halign: "left" },
    1: { cellWidth: 12, halign: "center" },
    2: { cellWidth: 12, halign: "center" },
    3: { cellWidth: 12, halign: "center" },
    4: { cellWidth: 12, halign: "center" },
    5: { cellWidth: 12, halign: "center" },
  };
  // Day columns auto width
  for (var i = 6; i < data.headers.length; i++) {
    colStyles[i] = { cellWidth: "auto", halign: "center" };
  }

  doc.autoTable({
    head: head,
    body: body,
    startY: 18,
    tableWidth: "auto",
    styles: {
      fontSize: 4.5,
      cellPadding: 0.8,
      overflow: "hidden",
      halign: "center",
    },
    columnStyles: colStyles,
    headStyles: {
      fillColor: [41, 128, 185],
      textColor: [255, 255, 255],
      fontSize: 4.5,
      fontStyle: "bold",
      halign: "center",
    },
    margin: { top: 18, left: 4, right: 4 },
    didParseCell: function(hookData) {
      if (hookData.section === "body") {
        var key = hookData.row.index + "_" + hookData.column.index;
        if (cellColorMap[key]) {
          hookData.cell.styles.fillColor = cellColorMap[key];
          hookData.cell.styles.textColor = [0, 0, 0];
        }
      }
    }
  });

  doc.save("leave_calendar_" + fromDate + "_to_" + toDate + ".pdf");
}

function applyThemeBasedOnUserSetting() {
  const deskTheme = document.getElementById("desk-theme")?.value;
  if (deskTheme === "D") { document.body.classList.add("dark-mode"); }
  else if (deskTheme === "L") { document.body.classList.remove("dark-mode"); }
  else if (deskTheme === "A") {
    const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
    if (prefersDark) { document.body.classList.add("dark-mode"); } else { document.body.classList.remove("dark-mode"); }
  }
}
applyThemeBasedOnUserSetting();
