import * as path from 'path'

import { Workbook } from './workbook'

const filePath = path.resolve('./test/test.xlsx')
const workbook = new Workbook(filePath)

workbook
  .init()
  .then(wb => {
    const headings = wb.worksheet('test').data().headings
    const testData = [
      { A: 'A2', B: 'B2', C: 'C2', D: 'C2' },
      { A: 'A3', B: 'B3', C: 'C3', D: 'C2' },
      { A: 'A4', B: 'B4', C: 'C4', D: 'C2' },
    ]
    return console.log(wb.convertJson(testData, headings))
  })
  .catch(error => Promise.reject(error))

export default Workbook

process.on('unhandledRejection', error => {
  console.log('caught error')
  console.error(error)
})
