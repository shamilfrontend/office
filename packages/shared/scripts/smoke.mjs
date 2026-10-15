#!/usr/bin/env node
/** Offline smoke for @office/shared — run after `npm run build`. */
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const require = createRequire(import.meta.url)
const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const s = require(join(root, 'dist/index.js'))

let failed = 0
function assert(cond, msg) {
  if (!cond) {
    console.error('FAIL:', msg)
    failed += 1
  } else {
    console.log('OK  ', msg)
  }
}

// 1) toolsForMode agent + buildingPlan
const agent = s.toolsForMode('agent', true, true, 'SHEET').map((t) => t.function.name)
assert(agent.includes('copy_formulas'), 'agent has copy_formulas')
assert(agent.includes('insert_image'), 'agent has insert_image')
assert(agent.includes('update_plan'), 'agent+buildingPlan has update_plan')
assert(agent.includes('set_formulas'), 'agent has set_formulas')
assert(!agent.includes('create_plan'), 'agent without plan mode has no create_plan')

const planMode = s.toolsForMode('plan', false, true, 'SHEET').map((t) => t.function.name)
assert(planMode.includes('create_plan'), 'plan mode has create_plan')
assert(!planMode.includes('set_cells'), 'plan mode has no write tools')

// 2) parsePlanMarkdown + hydrate (reload / «Сделать по плану»)
const md = `# Финансовый план

Обзор текста.

- [ ] Параметры
- [x] Шапка
- [ ] Формулы
`
const plan = s.parsePlanMarkdown(md)
assert(plan.name.includes('Финансовый'), 'plan name from heading')
assert(plan.todos.length === 3, '3 todos')
assert(plan.todos[0].id === '1' && !plan.todos[0].status, 'todo1 pending')
assert(plan.todos[1].id === '2' && plan.todos[1].status === 'completed', 'todo2 completed')
assert(!s.planAllTodosDone(plan), 'not all done')

// 3) update_plan merge by id (no «no plan» when hydrated)
const merged = s.mergePlanTodos(plan, [
  { id: '1', status: 'completed' },
  { id: '3', status: 'completed' },
])
assert(merged.todos[0].status === 'completed', 'merge by id todo1')
assert(merged.todos[2].status === 'completed', 'merge by id todo3')
assert(s.planAllTodosDone(merged), 'all todos done after merge')

// 4) update_plan merge by index when ids mismatch (agent vs markdown ids)
const byIndex = s.mergePlanTodos(
  { name: 'P', overview: '', todos: [{ id: 'a', content: 'x' }, { id: 'b', content: 'y' }] },
  [{ id: '1', status: 'completed' }, { id: '2', status: 'completed' }],
)
assert(byIndex.todos[0].status === 'completed', 'merge by index when ids differ')
assert(byIndex.todos[1].status === 'completed', 'merge by index second')

// 5) set_formulas empty holes skipped
const cells = s.collectFormulaCells(
  [['=G43'], ['=H43'], ['=F43'], [''], [''], ['=FV(1,2,3)']],
  45,
  1,
)
assert(cells.length === 4, 'skip empty formula cells')
assert(cells[3].formula === '=FV(1,2,3)' && cells[3].row === 50, 'last formula row offset')

// 6) copy_formulas relative adjust
const f1 = s.adjustFormulaRefs('=(C8+D8)*$B$4', 1, 0)
assert(f1 === '=(C9+D9)*$B$4', `relative row shift: ${f1}`)
const f2 = s.adjustFormulaRefs('=EDATE($B$5,0)', 5, 0)
assert(f2 === '=EDATE($B$5,0)', `absolute stays: ${f2}`)
const f3 = s.adjustFormulaRefs('=F8', 1, 0)
assert(f3 === '=F9', `simple shift: ${f3}`)

// 7) WRITE_TOOL_NAMES
assert(s.WRITE_TOOL_NAMES.has('copy_formulas'), 'WRITE has copy_formulas')
assert(s.WRITE_TOOL_NAMES.has('insert_image'), 'WRITE has insert_image')
assert(s.ALL_TOOL_NAMES.has('insert_image'), 'ALL has insert_image')

// 8) status normalize
assert(s.normalizePlanStatus('done') === 'completed', 'normalize done')
assert(s.normalizePlanStatus('COMPLETE') === 'completed', 'normalize COMPLETE')

if (failed) {
  console.error(`\n${failed} smoke check(s) failed`)
  process.exit(1)
}
console.log('\nAll shared smoke checks passed')
