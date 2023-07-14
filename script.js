let table = document.getElementsByClassName("sheet-body")[0],
rows = document.getElementsByClassName("rows")[0],
columns = document.getElementsByClassName("columns")[0]
tableExists = false

const generateTable = () => {
    const rowsNumber = +rows.value;
    const columnsNumber = +columns.value;
    table.innerHTML = ""
    for (let i = 1; i <= rowsNumber; i++) {
        let tableRow = '<tr></tr>'
        for (let j = 1; j <= columnsNumber; j++) {
            tableRow += '<td></td>'
        }
        table.innerHTML += tableRow
    }
    if(rowsNumber > 0 && columnsNumber > 0)
      tableExists =true
    else {
        swal("Oops", " Please enter valid numbers for rows and columns!", 'error')
    }
}

const ExportToExcel = (type, fn, dl) => {
    if(!tableExists){
        swal("Oops", " there is no generated table to be exported!", 'error')
        return
    }
    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
}


