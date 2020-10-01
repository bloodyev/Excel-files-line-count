
// run command
// node app

const log = console.log


const fs   = require('fs')
const path = require('path')
const XLSX = require('xlsx')

const stateArr = []

const getFiles = (dir, files_) => {

  files_ = files_ || []
  let files = fs.readdirSync(dir)
  for (let i in files) {
    let name = dir + '/' + files[i];
    if (fs.statSync(name).isDirectory()) {
      getFiles(name, files_)
    } else {
      files_.push(name)
    }
  }
  return files_

}

const factoryItems = (filename, lg) => {

  if (lg) {
    log(`\n=====================================================================\nFile load ${filename}\n=====================================================================\n`)
  }

  const workbook = XLSX.readFile(filename)

  // log(workbook)

  const items = workbook.Strings

  items.splice(0, 9)

  let key = 0
  let new_arr = []

  const next_fast = () => {
    key++
    run()
  }

  const next_sleep = () => {
    key++
    setTimeout(() => {
      run()
    }, randomInteger(1, 2) * 1000)
  }

  const run = () => {

    if (key != items.length) {

      const item = items[key]

      if (item.t) {
        if (item.t.indexOf('₽') != -1 || item.t.indexOf('БОНУС') != -1) {
          // log(`search ₽ no push: ${item.t}`)
          next_fast()
        } else {

          let i = item.t.split('x')
          let name = i[1].replace(/\r?\n/g, '')
          let name2 = name.replace(/\s+/g, ' ').trim()

          let count = parseInt(i[0])

          const obj = { name: name2, count: count }

          let indexOfKey = stateArr.findIndex(i => i.name === name2)

          // Не найдено
          if (indexOfKey === -1) {
            stateArr.push(obj)
          } else {
            stateArr[indexOfKey].count = stateArr[indexOfKey].count + obj.count
            if (lg) {
              log(`indexOfKey: [${indexOfKey}] - ${name2} | count: ${stateArr[indexOfKey].count}`)
            }
          }
          next_fast()
        }
      } else {
        next_fast()
      }

    } else {
      if (lg) {
        log(`\nFile exit parsing: ${filename}\n`)
      }
    }

  }

  run()

}

const start = () => {

  log('Loading........')

  const files = getFiles('./files')

  log(`Find ${files.length} files\n`)

  files.forEach((item, i) => {
    factoryItems(item, true);
  })

  // Интервао времени после которого выводим массив результата
  setTimeout(() => {
    log(stateArr)
  }, 5000)

}

const demo = () => {
  // factoryItems('filename', 'log - true:false')
  factoryItems('./files/514492_Zakaz_1_28_09.xlsx', true)
}

start()
