// Fetches the latest attachment from 4 separate ClickUp tasks (one task per
// source file - confirmed via the task ID pattern, e.g. "12yf9u-10935",
// which is a ClickUp Custom Task ID, not the default numeric task ID) and
// writes them into data/, ready for generate.js.
//
// Custom Task IDs need a different API call than plain task IDs: the
// request must include ?custom_task_ids=true&team_id={workspace_id}.
//
// Required env vars:
//   CLICKUP_TOKEN               - ClickUp API token (personal or app token)
//   CLICKUP_TEAM_ID              - workspace id, e.g. 36650298
//   CLICKUP_TASK_UEBERSICHT      - custom task id for Übersichtstabelle.xlsx
//   CLICKUP_TASK_GPSR_ISSUE      - custom task id for GPSR-Issue-Workbook
//   CLICKUP_TASK_GPSR_GRAPH      - custom task id for GPSR-Graph.xlsx
//   CLICKUP_TASK_PROHIBITED      - custom task id for Prohibited-Ingredients
const fs = require("fs");
const path = require("path");

const TOKEN = process.env.CLICKUP_TOKEN;
const TEAM_ID = process.env.CLICKUP_TEAM_ID;
const DATA_DIR = path.join(__dirname, "..", "data");

const FILE_TASKS = [
  { taskId: process.env.CLICKUP_TASK_UEBERSICHT, dest: "uebersicht.xlsx", label: "Übersichtstabelle" },
  { taskId: process.env.CLICKUP_TASK_GPSR_ISSUE, dest: "gpsr-issue.xlsx", label: "GPSR-Issue-Workbook" },
  { taskId: process.env.CLICKUP_TASK_GPSR_GRAPH, dest: "gpsr-graph.xlsx", label: "GPSR-Graph" },
  { taskId: process.env.CLICKUP_TASK_PROHIBITED, dest: "prohibited-ingredients.xlsx", label: "Prohibited-Ingredients" },
];

async function fetchTask(taskId) {
  const url = `https://api.clickup.com/api/v2/task/${taskId}?custom_task_ids=true&team_id=${TEAM_ID}`;
  const res = await fetch(url, { headers: { Authorization: TOKEN } });
  if (!res.ok) {
    throw new Error(`ClickUp API returned ${res.status} ${res.statusText} for task ${taskId}`);
  }
  return res.json();
}

async function main() {
  if (!TOKEN || !TEAM_ID) {
    console.error("Missing CLICKUP_TOKEN and/or CLICKUP_TEAM_ID environment variables.");
    process.exit(1);
  }
  const missingTaskIds = FILE_TASKS.filter((f) => !f.taskId).map((f) => f.label);
  if (missingTaskIds.length) {
    console.error(`Missing task id env var(s) for: ${missingTaskIds.join(", ")}`);
    process.exit(1);
  }

  fs.mkdirSync(DATA_DIR, { recursive: true });

  for (const file of FILE_TASKS) {
    const task = await fetchTask(file.taskId);
    const attachments = task.attachments || [];
    if (!attachments.length) {
      throw new Error(`Task ${file.taskId} (${file.label}) has no attachments.`);
    }
    // if a task has multiple attachments (re-uploads over time), the newest wins
    attachments.sort((a, b) => Number(b.date) - Number(a.date));
    const newest = attachments[0];
    const fileRes = await fetch(newest.url, { headers: { Authorization: TOKEN } });
    if (!fileRes.ok) {
      throw new Error(`Failed to download "${newest.title}" from task ${file.taskId}: ${fileRes.status} ${fileRes.statusText}`);
    }
    const buf = Buffer.from(await fileRes.arrayBuffer());
    fs.writeFileSync(path.join(DATA_DIR, file.dest), buf);
    console.log(`${file.dest} <- "${newest.title}" (Task ${file.taskId}, hochgeladen ${new Date(Number(newest.date)).toISOString()}, ${buf.length} bytes)`);
  }
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
