// One-off helper: calls bubblewrap's own project-generation logic directly,
// bypassing its CLI's interactive confirm prompt (which crashes in this
// non-TTY environment). Safe to delete after the Android project has been
// generated once — `bubblewrap build` works normally on subsequent runs
// since the checksum file this creates will already be present and valid.
const path = require('path');
const cliRoot = path.join(
  process.env.APPDATA, 'npm', 'node_modules', '@bubblewrap', 'cli', 'dist', 'lib'
);
const shared = require(path.join(cliRoot, 'cmds', 'shared.js'));

(async () => {
  const manifestFile = path.join(process.cwd(), 'twa-manifest.json');
  const targetDirectory = process.cwd();
  const ok = await shared.updateProject(true, null, undefined, targetDirectory, manifestFile);
  console.log('updateProject result:', ok);
  process.exit(ok ? 0 : 1);
})().catch((err) => {
  console.error('FAILED:', err);
  process.exit(1);
});
