const Koa = require('koa')

const app = new Koa()
const port = process.env.PORT || 3000
const host = process.env.HOST || 'http://localhost'

const bodyParser = require('koa-bodyparser')
const logger = require('koa-logger')
const convert = require('koa-convert')
const router = require('koa-better-router')().loadMethods()
const respond = require('koa-respond')

const chalk = require('chalk')
const rawBody = require('raw-body')

const fs = require('fs')
const crypto = require('crypto')

const RuleEngine = require('node-rules')
const instruct = require('./lib/instruct-o-matic')

const xlsxMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

const sheets = new Map([
  ['eligible_id', {
    timestamp: new Date(),
    data: fs.readFileSync('rules/eligible_id.xlsx')
  }]
])

const ruleEngine = new RuleEngine()
function promiseDecision (engine, param) {
  return new Promise(function (resolve, reject) {
    engine.execute(param, function (result) {
      resolve(result)
    })
  })
}

loadRules(sheets)
function loadRules (sheets) {
  for (let sheet of sheets.values()) {
    sheet.rules = instruct(sheet.data)
    ruleEngine.register(sheet.rules)
  }
  sheets.json = ruleEngine.toJSON()
  sheets.etag = etag(sheets.json)
}

function etag (value) {
  return crypto.createHash('md5').update(value).digest('hex')
}

app.use(convert(logger()))
app.use(respond())
app.use(bodyParser())

router
  .post('/decision', bodyParser(), async (ctx, next) => {
    switch (ctx.is('json')) {
      case 'json':
        let fact = ctx.request.body
        console.log(chalk.gray('Incoming fact:'), chalk.styles.yellow.open, fact, chalk.styles.yellow.close)
        try {
          fact = await promiseDecision(ruleEngine, fact)
          console.log(chalk.gray('Outgoing fact:'), chalk.styles.yellow.open, fact, chalk.styles.yellow.close)
          delete fact.result
          fact.rulesExecuted = fact.matchPath
          delete fact.matchPath
          fact.rulesETag = sheets.etag
          ctx.body = fact
          ctx.etag = etag(JSON.stringify({ rules: sheets.json, fact: fact }))
        } catch (err) {
          if (err instanceof TypeError) {
            // TODO fix message
            ctx.throw(406, 'Missing ' + err.message.match(/'.+'/)[0])
          } else {
            ctx.throw(err)
          }
        }
        break
      default:
        ctx.throw(406, 'json fact required')
    }
  })
  .get('/sheets', ctx => {
    ctx.body = Array.from(sheets.keys())
  })
  .get('/sheets/:id', ctx => {
    const id = ctx.params.id
    ctx.attachment(id + '.xlsx')
    ctx.type = xlsxMime
    const sheet = sheets.get(id)
    ctx.lastModified = sheet.timestamp
    ctx.body = sheet.data
    ctx.response.etag = etag(ctx.body)
  })
  .put('/sheets/:id', async ctx => {
    const id = ctx.params.id
    let data
    switch (ctx.is(xlsxMime, 'multipart/form-data')) {
      case xlsxMime:
        data = await rawBody(ctx.req)
        break
      case 'multipart/form-data':
        data = ctx.request.body
        break
      default:
        ctx.throw(406, `acceptable types are ${xlsxMime} and multipart/form-data`)
    }
    sheets.set(id, {
      data: data,
      timestamp: new Date(),
      rules: instruct(data)
    })
    loadRules(sheets)
    ctx.status = 201
  })

app.use(router.middleware())

if (!module.parent) app.listen(port, () => {
  console.log(chalk.magenta('Available routes:'))
  router.getRoutes().forEach(route => {
    console.log(`${chalk.yellow(route.method)}`,
      chalk.grey(`--> ${host}${port}`) + `${chalk.yellow(route.path)}`)
  })
  console.log(chalk.cyan('--------------------------------------------------------'))
  console.log(chalk.grey('Rules application webservice now listening at port:'), chalk.yellow(port))
  console.log(chalk.cyan('--------------------------------------------------------'))
  console.log(chalk.magenta('Traffic:\n'))
})
