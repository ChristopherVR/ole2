import { readLegacyOfficeMetadata, writeLegacyOfficeMetadata } from './vendor/ole2/index.js';

export function renderMetadata(bytes, fileName, report) {
	const section = document.querySelector('#metadata-result');
	const fields = document.querySelector('#metadata-fields');
	fields.replaceChildren();
	const properties = readLegacyOfficeMetadata(bytes);
	section.hidden = !properties || Object.keys(properties).length === 0;
	if (section.hidden) return;
	let current = bytes;
	for (const [key, value] of Object.entries(properties)) {
		const row = document.createElement('div');
		row.className = 'metadata-row';
		const label = document.createElement('label');
		label.textContent = key;
		const input = document.createElement('input');
		input.type = 'text'; input.value = value;
		input.id = `metadata-${key}`; label.htmlFor = input.id;
		const button = document.createElement('button');
		button.type = 'button'; button.textContent = 'Apply';
		button.addEventListener('click', () => {
			const updated = writeLegacyOfficeMetadata(current, key, input.value);
			if (updated === current) {
				report(input.value === readLegacyOfficeMetadata(current)?.[key]
					? 'This property is unchanged.'
					: 'This edit cannot fit safely in the existing property. Try a shorter value in the original character set.', 'info');
				return;
			}
			current = updated;
			report(`Updated ${key}. Download the edited file to keep your changes.`, 'success');
		});
		row.append(label, input, button); fields.append(row);
	}
	document.querySelector('#download-metadata').onclick = () => {
		const url = URL.createObjectURL(new Blob([current], { type: 'application/octet-stream' }));
		const link = document.createElement('a');
		link.href = url; link.download = `edited-${fileName}`;
		link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
	};
}
