import * as XLSX from 'xlsx'
import * as path from 'path'
import * as fs from 'fs'

const inputFilePath = 'D:/folder/file.xlsx'
const chunkSize = 20000
const fieldDate = ['fecApertura']

function convertDates(data: Record<string, any>[]): Record<string, any>[] {
  return data.map((row) => {
    const newRow: Record<string, any> = { ...row }
    for (const field of fieldDate) {
      if (newRow[field] && typeof newRow[field] === 'number') {
        const date = XLSX.SSF.parse_date_code(newRow[field])
        if (date) {
          const year = date.y.toString().padStart(4, '0')
          const month = date.m.toString().padStart(2, '0')
          const day = date.d.toString().padStart(2, '0')
          newRow[field] = `${year}-${month}-${day}`
        }
      }
    }
    return newRow
  })
}

function divideExcel(filePath: string, rowsByFile: number): void {
  if (!fs.existsSync(filePath)) {
    console.error(`❌ File ${filePath} not exists`)
    return
  }

  const workbook = XLSX.readFile(filePath, { cellDates: true })
  const sheetName = workbook.SheetNames[0]
  const worksheet = workbook.Sheets[sheetName]
  const data: Record<string, any>[] = XLSX.utils.sheet_to_json(worksheet, {
    defval: '',
  })

  console.log(`Total rows: ${data.length}`)

  const fileName = path.basename(filePath, path.extname(filePath))
  const outputDir = path.join(path.dirname(filePath), `${fileName}_partes`)

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir)
  }

  const totalParts = Math.ceil(data.length / rowsByFile)

  for (let i = 0; i < totalParts; i++) {
    const init = i * rowsByFile
    const end = init + rowsByFile
    const chunk = convertDates(data.slice(init, end))

    const newWorkbook = XLSX.utils.book_new()
    const newWorksheet = XLSX.utils.json_to_sheet(chunk)

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1')

    const outputFile = path.join(outputDir, `${fileName}_part_${i + 1}.xlsx`)
    XLSX.writeFile(newWorkbook, outputFile)

    console.log(`✅ File generate: ${outputFile}`)
  }

  console.log('🎉 Finish Successfully!')
}

divideExcel(inputFilePath, chunkSize)
