---
trigger: always_on
description: Always sync and maintain project context in .agent-memory/ across PCs
---

# Cross-PC Agent Memory Sync Rule

At the start of every session, always read `.agent-memory/PROJECT_STATE.md` and `.agent-memory/TASKS.md` first. Whenever we complete a task, change architecture, or before ending a session, automatically update the files in `.agent-memory/` so the other PC can pick up seamlessly.
