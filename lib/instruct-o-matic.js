const debug = require('debug')('instruct-o-matic')
const XLSX = require('xlsx')

const decodeCell = XLSX.utils.decode_cell

function getGapAndBindings (worksheet) {
  const range = worksheet['!ref']
  delete worksheet['!ref']

  let needGap = true
  let onBindingRow = false

  const result = {
    bindings: {}
  }
  let checkCell = range.substring(0, range.indexOf(':'))
  for (const cell in worksheet) {
    let decoded = decodeCell(cell)
    if (cell === checkCell) {
      delete worksheet[cell]
      continue
    }
    if (needGap) {
      if (decoded.c !== decodeCell(checkCell).c + 1) {
        result.gapColumn = decodeCell(checkCell).c + 1
        needGap = false
        debug('Found gap:', cell)
      } else {
        checkCell = cell
      }
    } else {
      if (onBindingRow) {
        if (decodeCell(checkCell).r === decoded.r) {
          result.bindings[decoded.c] = worksheet[cell].v
        } else {
          break
        }
      } else {
        if (worksheet[cell].v === 'BINDING') {
          onBindingRow = true
          checkCell = cell
          debug('Found bindings:', cell)
        }
      }
    }
    delete worksheet[cell]
  }
  debug('Found table data', result)
  return result
}

function getRuleValues (worksheet, gapColumn) {
  const result = {}
  let currentRow
  let rule

  for (const cell in worksheet) {
    const decoded = decodeCell(cell)
    if (!currentRow || decoded.r !== currentRow) {
      currentRow = decoded.r
      rule = {
        conditions: {},
        consequences: {}
      }
      result[worksheet[cell].v] = rule
      debug('Rule name:', worksheet[cell].v)
    } else {
      if (decoded.c < gapColumn) {
        debug('Condition(s):', worksheet[cell].v)
        rule.conditions[decoded.c] = worksheet[cell].v
      } else {
        debug('Consequence:', worksheet[cell].v)
        rule.consequences[decoded.c] =
          (worksheet[cell].v === 'TRUE' || worksheet[cell].v === 'FALSE')
          ? worksheet[cell].v.toLowerCase()
          : worksheet[cell].v
      }
    }
  }
  return result
}

function getRules (bindings, values) {
  const result = []

  for (const ruleName in values) {
    const conditionString = bind(values[ruleName].conditions, bindings, ' && ')
    debug('Condition for %s: %o', ruleName, conditionString)
    const consequenceString = bind(values[ruleName].consequences, bindings, '\n')
    debug('Consequence for %s: %o', ruleName, consequenceString)

    const rule = {
      name: ruleName,
      condition: new Function('R', 'R.when(' + conditionString + ')'),
      consequence: new Function('R', consequenceString + '\nR.stop()')
    }
    result.push(rule)
  }

  return result

  function bind (values, bindings, separator) {
    let result = ''
    for (const column in values) {
      if (result.length > 0) {
        result += separator
      }
      result += bindings[column].replace('$value', values[column])
    }
    return result
  }
}

module.exports = function (sheet) {
  const workbook = XLSX.read(sheet)
  const worksheet = workbook.Sheets[workbook.SheetNames[0]]
  const { gapColumn, bindings } = getGapAndBindings(worksheet)
  const ruleValues = getRuleValues(worksheet, gapColumn)
  const rules = getRules(bindings, ruleValues)
  return rules
}
