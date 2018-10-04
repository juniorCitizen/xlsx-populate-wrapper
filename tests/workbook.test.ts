import * as fs from 'fs-extra'
import * as path from 'path'

import WsData from '../src/wsData'
import {IWsDataClass, IWsDataStructure} from '../src/wsData'

describe('testing the WsData class', () => {
  test('the class to be defined', () => {
    expect(WsData).toBeDefined()
  })

  describe('basic instantiation', () => {
    let wsData: IWsDataClass
    beforeAll(() => {
      wsData = new WsData()
    })
    test('to receive an instance when instantiated', () => {
      expect(wsData).toBeInstanceOf(WsData)
    })

    test('a getData method to be defined', () => {
      expect(wsData.getData()).toBeDefined()
    })

    test('a getHeadings method to be defined', () => {
      expect(wsData.getHeadings()).toBeDefined()
    })

    test('getAoaData method to be defined', () => {
      expect(wsData.getAoaData()).toBeDefined()
    })

    test('getJsonData method to be defined', () => {
      expect(wsData.getJsonData()).toBeDefined()
    })
  })

  describe('checking class members are initialized correctly', () => {
    test('to receive empty class members if constructor params are missing', () => {
      const initData: IWsDataStructure = {
        headings: [],
        aoaData: [],
        jsonData: [],
      }
      const wsData: IWsDataClass = new WsData(initData)
      const data: IWsDataStructure = wsData.getData()
      const a: boolean = data.headings.length === 0
      const b: boolean = data.aoaData.length === 0
      const c: boolean = data.jsonData.length === 0
      expect(a && b && c).toBeTruthy()
    })

    describe('members to be initialized correctly with params', () => {
      const filePath: string = path.resolve('./tests/init.json')
      let wsData: IWsDataClass
      let initData: IWsDataStructure
      beforeAll(async () => {
        initData = await fs.readJson(filePath)
        wsData = new WsData(initData)
      })
      test('headings length to be correct', async () => {
        expect(wsData.getHeadings().length).toBe(initData.headings.length)
        console.log(wsData.getHeadings())
      })
      test('aoaData length to be correct', async () => {
        expect(wsData.getAoaData().length).toBe(initData.aoaData.length)
        console.log(wsData.getAoaData())
      })
      test('jsonData length to be correct', async () => {
        expect(wsData.getJsonData().length).toBe(initData.jsonData.length)
        console.log(wsData.getJsonData())
      })
    })
  })
})

// import * as path from 'path'

// import {Workbook} from '../src/workbook'
// import {IWorksheetData} from '../src/worksheet'

// const mockFilePath = path.resolve('./tests/test.xlsx')

// describe('testing the Workbook class', () => {
//   test('to throw while retriving data from uninitialized instance', () => {
//     function attempt() {
//       const wb = new Workbook(mockFilePath)
//       return wb.data()
//     }
//     expect(attempt).toThrow('workbook is not ready')
//   })

//   test('to get an instance of object after init', async () => {
//     const wb = new Workbook(mockFilePath)
//     const instance = await wb.initialize()
//     expect(instance).toBeInstanceOf(Workbook)
//   })

//   describe('test data I/O operations', () => {
//     let workbook: Workbook
//     let workingData: IWorksheetData[]
//     beforeAll(async () => {
//       try {
//         workbook = new Workbook(mockFilePath)
//         await workbook.initialize()
//       } catch (error) {
//         throw error
//       }
//     })
//     describe('test writing operations', () => {
//       const data = {
//         headings: [
//           '1st row heading',
//           'extra row (should appear)',
//           '第二行抬頭',
//           '3rd row heading',
//           'extra row (should appear at the end)',
//         ],
//         aoaData: [
//           ['a', 2, '第三'],
//           ['a', '', null],
//           ['a', null, undefined],
//           ['a', undefined, ''],
//           [],
//           [null, null, null],
//           [null, 'b', 'c'],
//           [undefined, undefined, undefined],
//           [undefined, 'b', 'c'],
//           ['', '', ''],
//           ['', 'b', 'c'],
//         ],
//         jsonData: [
//           {
//             '1st row heading': 'a',
//             第二行抬頭: 2,
//             '3rd row heading': '第三',
//           },
//           {
//             '1st row heading': 'a',
//             第二行抬頭: '',
//             '3rd row heading': null,
//           },
//           {
//             '1st row heading': 'a',
//             第二行抬頭: null,
//             '3rd row heading': undefined,
//           },
//           {
//             '1st row heading': 'a',
//             第二行抬頭: undefined,
//             '3rd row heading': '',
//           },
//           {},
//           {
//             '1st row heading': null,
//             第二行抬頭: null,
//             '3rd row heading': null,
//           },
//           {
//             '1st row heading': null,
//             第二行抬頭: 'b',
//             '3rd row heading': 'c',
//           },
//           {
//             '1st row heading': undefined,
//             第二行抬頭: undefined,
//             '3rd row heading': undefined,
//           },
//           {
//             '1st row heading': undefined,
//             第二行抬頭: 'b',
//             '3rd row heading': 'c',
//           },
//           {
//             '1st row heading': '',
//             第二行抬頭: '',
//             '3rd row heading': '',
//           },
//           {
//             '1st row heading': '',
//             第二行抬頭: 'b',
//             '3rd row heading': 'c',
//           },
//           {
//             '1st row heading': 'a',
//             第二行抬頭: undefined,
//             '3rd row heading': '',
//             'extra row (should not appear)': '',
//           },
//         ],
//       }
//       test('writing to 4th sheet complete without errors', async () => {
//         expect.assertions(1)
//         await workbook.update('sheet3', data)
//         workingData = workbook.data()
//       })
//     })
//     describe('test reading operations', () => {
//       test('workingData to be an array of correct length', () => {
//         expect(workingData.length).toBe(4)
//       })
//       test('1st sheet data headings to be of correct length', () => {
//         expect(workingData[0].headings.length).toBe(8)
//       })
//       test('2nd sheet data aoaData to be of correct length', () => {
//         expect(workingData[1].aoaData.length).toBe(5)
//       })
//       test('3rd sheet to be of the correct name', () => {
//         expect(workbook.worksheetNames()[2]).toBe('sheet2')
//       })
//     })
//     // test('headings of the 5th sheet(blank) to have 0 as length', () => {
//     //   expect(workingData[4].headings.length).toBe(0)
//     // })
//     // test('aoaData of the 5th sheet(blank) to have 0 as length', () => {
//     //   expect(workingData[4].aoaData.length).toBe(0)
//     // })

//     // test('jsonData of the 5th sheet(blank) to have 0 as length', () => {
//     //   expect(workingData[4].jsonData.length).toBe(0)
//     // })
//   })
// })
