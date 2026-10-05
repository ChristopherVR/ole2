const docs = Bun.spawnSync(['bunx', 'vitepress', 'build', 'docs'], { stdout: 'inherit', stderr: 'inherit' });
if (docs.exitCode !== 0) throw new Error('VitePress build failed.');
