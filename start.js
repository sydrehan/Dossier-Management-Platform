import { spawn } from "child_process";

console.log("=========================================");
console.log("Starting Document Merging POC Services...");
console.log("=========================================");

const backend = spawn("npm", ["start"], {
  cwd: "./backend",
  shell: true,
  stdio: "inherit",
});

const frontend = spawn("npm", ["run", "dev"], {
  cwd: "./frontend",
  shell: true,
  stdio: "inherit",
});

function cleanup() {
  console.log("\nStopping all services...");
  try {
    backend.kill();
  } catch (e) {}
  try {
    frontend.kill();
  } catch (e) {}
  process.exit();
}

process.on("SIGINT", cleanup);
process.on("SIGTERM", cleanup);
