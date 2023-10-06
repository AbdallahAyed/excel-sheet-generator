let table = document.getElementsByClassName("sheet-body")[0],
  rows = document.getElementsByClassName("rows")[0],
  columns = document.getElementsByClassName("columns")[0];

let tableExists = false;

const generateTable = () => {
  let rowsNumber = parseInt(rows.value),
    columnsNumber = parseInt(columns.value);

  console.log(rowsNumber, columnsNumber);

  table.innerHTML = "";

  for (let i = 0; i < rowsNumber; i++) {
    let tableRow = "";

    for (let j = 0; j < columnsNumber; j++) {
      tableRow += `<td contenteditable></td>`;
    }
    table.innerHTML += tableRow;
  }

  if (rowsNumber > 0 && columnsNumber > 0) {
    tableExists = true;
  }

  if (isNaN(rowsNumber) || isNaN(columnsNumber)) {
    Swal.fire("Fields are Empty", "check the fields?", "question");
  }
};

const ExportToExcel = (type, fn, dl) => {
  if (!tableExists) {
    Swal.fire({
      icon: "error",
      title: "Oops...",
      text: "Can't export before generating",
    });
    return;
  }
  let elt = table;
  let wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
  return dl
    ? XLSX.write(wb, { bookType: type, bookSST: true, type: "base64" })
    : XLSX.writeFile(wb, fn || "MyNewSheet." + (type || "xlsx"));
};
