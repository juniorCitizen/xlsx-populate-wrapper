import { Workbook } from './workbook'

export default Workbook

process.on('unhandledRejection', error => {
  throw error
})

process.on('uncaughtException', error => {
  throw error
})
