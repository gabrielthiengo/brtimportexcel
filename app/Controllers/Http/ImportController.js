'use strict'

const Helpers = use('Helpers')
const Excel = use('exceljs')

class ImportController {
  async store({ request, response }) {
    const upload = request.file('file', { size: '5mb' })

    const fileName = `${Date.now()}.xlsx`

    await upload.move(Helpers.tmpPath('uploads'), {
      name: fileName
    })

    if(!upload.moved()) {
      return response.status(500).send({ error: 'Houve um erro ao importar' })
    }

    let workbook = new Excel.Workbook()

    workbook = await workbook.xlsx.readFile(`temp/uploads/${fileName}`)

    const explanation = workbook.getWorksheet('Mailing')
    console.log(explanation)

    console.log(explanation.getCell('A' + 2).fill)

    return [{
        column: explanation.getCell('A' + 1).value,
        value: explanation.getCell('A' + 2).value,
        pattern: explanation.getCell('A' + 2).fill,
        color: explanation.getCell('A' + 2).fill.fgColor.argb
      },
      {
        column: explanation.getCell('A' + 1).value,
        value: explanation.getCell('A' + 3).value,
        pattern: explanation.getCell('A' + 3).fill,
        color: explanation.getCell('A' + 3).fill.fgColor.argb
      }]
  }
}

module.exports = ImportController
