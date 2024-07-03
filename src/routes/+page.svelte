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
	let contracts = contracts_json.reinsurance_contracts;
	let errorMessages = [];

	async function handleFileUpload(event) {
		console.log('File upload started');
		const file = event.target.files[0];
		if (!file) {
			console.error('No file selected');
			errorMessages = ['No file selected'];
			return;
		}
		console.log('File selected:', file.name);

		const workbook = new Workbook();
		errorMessages = []; // Reset error messages

		try {
			console.log('Loading file...');
			await workbook.xlsx.load(file);
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
			console.error('Error in handleFileUpload:', error);
			errorMessages = [
				...errorMessages,
				`Error reading file: ${error.message}`
			];
		}
	}

	$: console.log('Contracts updated:', contracts);
	$: console.log('Error messages:', errorMessages);
</script>

<head>
	<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@picocss/pico@2/css/pico.min.css" />
</head>

<h1>Welcome to Green Owl</h1>
<p>Risk assessment for reinsurance portfolios</p>

<input type="file" accept=".xlsx, .xls" on:change={handleFileUpload} />
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
