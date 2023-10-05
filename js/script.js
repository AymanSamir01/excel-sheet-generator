let table = document.getElementsByClassName("sheet-body")[0],
  rows = document.getElementsByClassName("rows")[0],
  columns = document.getElementsByClassName("columns")[0];
tableExists = false;

const generateTable = () => {
  let rowsNumber = parseInt(rows.value),
    columnsNumber = parseInt(columns.value);
  table.innerHTML = "";
  for (let i = 0; i < rowsNumber; i++) {
    var tableRow = "";
    for (let j = 0; j < columnsNumber; j++) {
      tableRow += `<td contenteditable></td>`;
    }
    table.innerHTML += tableRow;
  }
  if (rowsNumber > 0 && columnsNumber > 0) {
    tableExists = true;
    Swal.fire({
      position: "top-end",
      icon: "success",
      title: "table generated successfully",
      showConfirmButton: false,
      timer: 1000,
    });
  } else {
    Swal.fire({
      icon: "error",
      position: "center",
      title: "All fields are required",
      confirmButtonText: "ok",
    });
  }
};

const ExportToExcel = (type, fn, dl) => {
  if (!tableExists) {
    Swal.fire({
      icon: "error",
      position: "center",
      title: "there is no generated table to be exported",
      confirmButtonText: "ok",
    });
    return;
  }
  var elt = table;
  var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
  Swal.fire({
    position: "top-end",
    icon: "success",
    title: "table exported successfully",
    showConfirmButton: false,
    timer: 1000,
  });
  return dl
    ? XLSX.write(wb, { bookType: type, bookSST: true, type: "base64" })
    : XLSX.writeFile(wb, fn || "MyNewSheet." + (type || "xlsx"));
};
