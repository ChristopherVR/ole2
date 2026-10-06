import { defineConfig } from 'vitepress';

export default defineConfig({
	title: 'ole2',
	description: 'Parse, edit and serialize legacy Office binary files (DOC, XLS, PPT, VSD) and OLE2 compound files in TypeScript.',
	base: '/ole2/',
	cleanUrls: true,
	themeConfig: {
		nav: [
			{ text: 'Guide', link: '/getting-started' },
			{ text: 'Formats', link: '/formats/models' },
			{ text: 'API', link: '/api' },
			{ text: 'Limitations', link: '/limitations' },
			{ text: 'Releases', link: '/releases/0.12.0' },
		],
		sidebar: [
			{
				text: 'Guide',
				items: [
					{ text: 'Getting started', link: '/getting-started' },
					{ text: 'Limitations', link: '/limitations' },
				],
			},
			{
				text: 'Formats',
				items: [
					{ text: 'Document models', link: '/formats/models' },
					{ text: 'Word (DOC)', link: '/formats/word' },
					{ text: 'Excel (XLS)', link: '/formats/excel' },
					{ text: 'PowerPoint (PPT)', link: '/formats/powerpoint' },
					{ text: 'Visio (VSD)', link: '/formats/visio' },
					{ text: 'Publisher (PUB)', link: '/formats/publisher' },
					{ text: 'OLE2 containers', link: '/formats/containers' },
				],
			},
			{
				text: 'Reference',
				items: [
					{ text: 'API', link: '/api' },
					{ text: 'PPT export', link: '/ppt-export' },
					{ text: 'Serialization contract', link: '/document-model' },
					{ text: 'Format coverage', link: '/legacy-coverage' },
					{ text: 'CFB compatibility', link: '/cfb-compatibility' },
					{ text: 'CFB v4 preservation', link: '/cfb-v4-preservation' },
				],
			},
			{
				text: 'Validation notes',
				collapsed: true,
				items: [
					{ text: 'DOC paragraph alignment', link: '/doc-paragraph-alignment' },
					{ text: 'DOC underline', link: '/doc-underline' },
					{ text: 'PPT character runs', link: '/ppt-character-runs' },
					{ text: 'PPT shape native cases', link: '/ppt-shape-native-cases' },
					{ text: 'VSD shape order', link: '/vsd-shape-order' },
					{ text: 'Visio native validation', link: '/visio-com-native-validation' },
					{ text: 'XLS preserved string edits', link: '/xls-preserved-string-edits' },
				],
			},
			{
				text: 'Release notes',
				collapsed: true,
				items: ['0.12.0', '0.11.0', '0.9.0', '0.8.0', '0.7.0', '0.6.0', '0.5.0'].map((v) => ({
					text: v,
					link: `/releases/${v}`,
				})),
			},
		],
		socialLinks: [
			{ icon: 'github', link: 'https://github.com/ChristopherVR/ole2' },
			{ icon: 'npm', link: 'https://www.npmjs.com/package/@christophervr/ole2' },
		],
		search: { provider: 'local' },
		editLink: {
			pattern: 'https://github.com/ChristopherVR/ole2/edit/main/docs/:path',
		},
		footer: {
			message: 'Released under the Apache-2.0 License.',
		},
	},
});
