const { existsSync } = require("fs");
const { join } = require("path");
const { spawnSync } = require("child_process");

const rootStandaloneServer = join(process.cwd(), "server.js");
const localStandaloneServer = join(process.cwd(), ".next", "standalone", "server.js");

if (existsSync(rootStandaloneServer)) {
  require(rootStandaloneServer);
} else if (existsSync(localStandaloneServer)) {
  require(localStandaloneServer);
} else {
  const nextBin = process.platform === "win32" ? "next.cmd" : "next";
  const result = spawnSync(nextBin, ["start"], {
    stdio: "inherit",
    shell: process.platform === "win32",
  });

  process.exit(result.status ?? 1);
}
