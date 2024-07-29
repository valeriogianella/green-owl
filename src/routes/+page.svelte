<script>
	import pkg from 'exceljs';
	const { Workbook } = pkg;
	const headers = [
		'Contract ID',
		'Name',
		'Line of Business',
		'Country',
		'Currency',
		'Expected Loss',
		'Standard Deviation'
	];
	import contracts_json from './contracts.json';
	let fileHandle = null;
	let contracts = [];
	let errorMessages = [];
	let fileName = '';

	async function handleFileOpen() {
		try {
			[fileHandle] = await window.showOpenFilePicker({
				types: [
					{
						description: 'Excel Files',
						accept: {
							'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
							'application/vnd.ms-excel': ['.xls']
						}
					}
				]
			});
			fileName = fileHandle.name;
			await readFile();
		} catch (error) {
			console.error('Error in handleFileOpen:', error);
			errorMessages = [...errorMessages, `Error reading file: ${error.message}`];
		}
	}

	async function readFile() {
		if (!fileHandle) {
			console.error('No file handle available');
			errorMessages = ['No file selected. Please upload a file first.'];
			return;
		}

		const file = await fileHandle.getFile();
		const workbook = new Workbook();
		errorMessages = []; // Reset error messages

		try {
			console.log('Loading file...');
			const arrayBuffer = await file.arrayBuffer();
			await workbook.xlsx.load(arrayBuffer);
			console.log('File loaded successfully');

			const worksheet = workbook.worksheets[0];
			if (!worksheet) {
				throw new Error('No worksheet found in the file');
			}

			let headerValid = true;
			console.log('Checking headers...');
			headers.forEach((expectedHeader, index) => {
				const cell = worksheet.getRow(1).getCell(index + 1);
				const actualHeader = cell.value || '';
				if (actualHeader !== expectedHeader) {
					errorMessages = [
						...errorMessages,
						`Invalid header in column ${index + 1}: expected <i>${expectedHeader}</i> but received <i>${actualHeader === '' ? '(empty)' : actualHeader}</i>.`
					];
					headerValid = false;
				}
			});

			if (!headerValid) {
				console.log('Invalid headers found');
				return; // Exit the function if headers are invalid
			}

			console.log('Headers are valid, processing data...');
			const jsonData = [];
			worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
				if (rowNumber > 1) {
					jsonData.push({
						contract_id: row.getCell(1).value,
						name: row.getCell(2).value,
						line_of_business: row.getCell(3).value,
						country: row.getCell(4).value,
						currency: row.getCell(5).value,
						expected_loss: row.getCell(6).value,
						standard_deviation: row.getCell(7).value
					});
				}
			});

			console.log('Data processed, updating contracts');
			contracts = jsonData;
			console.log('Contracts updated');
		} catch (error) {
			console.error('Error in readFile:', error);
			errorMessages = [...errorMessages, `Error reading file: ${error.message}`];
		}
	}

	async function refreshData() {
		if (fileHandle) {
			await readFile();
		} else {
			errorMessages = ['No file selected. Please upload a file first.'];
		}
	}
</script>

<head>
	<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@picocss/pico@2/css/pico.min.css" />
</head>

<h1>Welcome to Green Owl</h1>
<p>Risk assessment for reinsurance portfolios</p>
<div style="display: flex; gap: 10px;">
	<button on:click={handleFileOpen}>Open File</button>

	{#if fileName}
		<p>Current file: {fileName}</p>
		<button on:click={refreshData}>Refresh Data</button>
	{/if}
</div>

{#if errorMessages.length > 0}
	<div class="error">
		{#each errorMessages as error}
			<p>{@html error}</p>
		{/each}
	</div>
{/if}

<div>
	<table>
		<thead>
			<tr>
				{#each headers as header}
					<th>{header}</th>
				{/each}
			</tr>
		</thead>
		<tbody>
			{#each contracts as contract}
				<tr>
					{#each Object.values(contract) as value}
						<td>{value}</td>
					{/each}
				</tr>
			{/each}
		</tbody>
	</table>
</div>
