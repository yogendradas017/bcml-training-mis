---
description: 4-agent squad — Architect → Engineer → Reviewer → Optimizer pass on a task
---

User wants a 4-role pipeline on TMS task. This is a Workflow tool task — orchestrate 4 subagents.

## Phases

1. **Architect** — design approach. Output: 3-step plan, files to touch, risks.
2. **Engineer** — implement plan. Output: actual edits.
3. **Reviewer** — senior code review of engineer's output. Output: findings list (bugs, missing edge cases, CLAUDE.md violations).
4. **Optimizer** — apply reviewer's fixes + tighten perf/clarity. Output: final diff.

## How to run
Use the Workflow tool. Pipeline pattern, each item flows through all 4 stages.

```js
export const meta = {
  name: 'tms-squad',
  description: 'Architect → Engineer → Reviewer → Optimizer pass',
  phases: [
    {title:'Architect'}, {title:'Engineer'}, {title:'Reviewer'}, {title:'Optimizer'}
  ],
}
const task = args || 'unspecified — ask user'
const arch = await agent(`Architect: ${task}. Output 3-step plan + files + risks.`, {phase:'Architect'})
const eng  = await agent(`Engineer: implement this plan. Edit files. Plan:\n${arch}`, {phase:'Engineer'})
const rev  = await agent(`Senior review of these edits. Findings only — file:line + severity. CLAUDE.md rules apply.\n${eng}`, {phase:'Reviewer'})
const opt  = await agent(`Apply reviewer findings + tighten. Output final diff.\nReview:\n${rev}\nOriginal edits:\n${eng}`, {phase:'Optimizer'})
return {arch, eng, rev, opt}
```

## Rules
- Reviewer must NOT rubber-stamp. If clean, say "no issues, here's why" with evidence.
- Optimizer keeps behavior identical; only quality/perf changes.
- Final output: 4 sections collapsed, recommend commit message.

$ARGUMENTS  # the task to run the squad on
