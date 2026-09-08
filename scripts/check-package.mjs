import { spawnSync } from 'node:child_process'

const result = spawnSync(
  'npm',
  ['pack', '--dry-run', '--json', '--ignore-scripts'],
  { encoding: 'utf8' },
)

if (result.status !== 0) {
  throw new Error(result.stderr || 'npm pack dry-run failed')
}

const jsonStart = result.stdout.search(/^\[/m)
if (jsonStart < 0) {
  throw new Error(`npm pack returned no JSON manifest:\n${result.stdout}`)
}
const [manifest] = JSON.parse(result.stdout.slice(jsonStart))
const paths = manifest.files.map(file => file.path)
const allowed = path =>
  path.startsWith('build/') ||
  path.startsWith('docs/') ||
  ['AGENTS.md', 'LICENSE', 'README.md', 'package.json'].includes(path)

const unexpected = paths.filter(path => !allowed(path))
if (unexpected.length > 0) {
  throw new Error(`Unexpected packaged files: ${unexpected.join(', ')}`)
}
if (!paths.includes('build/cli.js')) {
  throw new Error('Published package is missing build/cli.js')
}
if (paths.some(path => path.endsWith('.map'))) {
  throw new Error('Published package must not contain source maps')
}

process.stdout.write(
  `Tarball allowlist OK: ${paths.length} files, ${manifest.size} bytes\n`,
)
