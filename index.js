const puppeteer = require('puppeteer');
const CREDS = require('./creds.json');
// Require library
const xl = require('excel4node');

async function cleanNumbers(dirty) {
	var result = dirty.replace('.', '').replace(',', '.');
	return Number(result)
}

async function getData() {
	// the step indicates which round is played in the game
	const STEP = 10;

	// Create a new instance of a Workbook class
	var wb = new xl.Workbook();

	// #################################### Login ####################################
	const USERNAME_SELECTOR = '#portlet-wrapper-58 > div.portlet-content > div > div > form > fieldset > div:nth-child(1) > input[type="text"]';
	const PASSWORD_SELECTOR = '#_58_password';
	const BUTTON_SELECTOR = '#portlet-wrapper-58 > div.portlet-content > div > div > form > fieldset > div.button-holder > input[type="submit"]';

	const browser = await puppeteer.launch({headless: false});
	const page = await browser.newPage();
	await page.goto('http://tum.go4c.org/web/guest/home');

	await page.evaluate(function () {
		document.querySelector('#portlet-wrapper-58 > div.portlet-content > div > div > form > fieldset > div:nth-child(1) > input[type="text"]').value = '';
	});
	await page.click(USERNAME_SELECTOR);
	await page.keyboard.type(CREDS.username);

	await page.evaluate(function () {
		document.querySelector('#portlet-wrapper-58 > div.portlet-content > div > div > form > fieldset > div.button-holder > input[type="submit"]').value = '';
	});

	await page.click(PASSWORD_SELECTOR);
	await page.keyboard.type(CREDS.password);
	// submit login
	await page.click(BUTTON_SELECTOR);
	await page.waitForNavigation();
	await page.screenshot({path: 'test.png'});

	// #################################### Balance Sheet ####################################
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/management-overview');
	await page.evaluate(() => {
		var result = {
			rows: [],
		};

		// Left Side
		var titles = document.querySelectorAll('#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > table > tbody > tr > td:nth-child(1) > table > tbody > tr > td:nth-child(1)');
		//console.log(titles);
		var datasets = document.querySelectorAll('#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > table > tbody > tr > td:nth-child(1) > table > tbody > tr > td:nth-child(2)');
		result.rows.push({
			key: titles[0] && titles[0].innerText,
			value: ''
		});
		var i = 0;
		while (i < datasets.length) {

			var title = (titles[i + 1] && titles[i + 1].innerText) || ' ';
			var val = (datasets[i] && datasets[i].innerText) || ' ';

			result.rows.push({
				key: title,
				value: val.replace('.', '').replace(',', '.')
			});
			i++;
		}
		// Right Side
		titles = document.querySelectorAll('#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > table > tbody > tr > td:nth-child(2) > table > tbody > tr > td:nth-child(1)');
		datasets = document.querySelectorAll('#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > table > tbody > tr > td:nth-child(2) > table > tbody > tr > td:nth-child(2)');
		result.rows.push({
			key: titles[0] && titles[0].innerText,
			value: ''
		});
		i = 0;
		while (i < datasets.length) {
			title = (titles[i + 1] && titles[i + 1].innerText) || ' ';
			val = (datasets[i] && datasets[i].innerText) || ' ';

			result.rows.push({
				key: title,
				value: val.replace(/\./g, '').replace(',', '.')
			});
			i++;
		}
		return Promise.resolve(result);
	}).then(function (result) {
		var balanceSheet = wb.addWorksheet('Balance Sheet');
		// Fix the new formatting based on the actual round stuff
        if(STEP > 1 && result.rows && result.rows.length>15){
            result.rows.splice(1,1);
            result.rows.splice(14,1);
        }
		for (var i = 0; i < result.rows.length; i++) {
			// Write title to column A
			balanceSheet.cell(1 + i, 1).string(result.rows[i].key);
			// Check if its a number
			// write results in column B+STEP
			if (result.rows[i].value === ''|| isNaN(result.rows[i].value)) {

			} else {
				balanceSheet.cell(1 + i, 1 + STEP).number(Number(result.rows[i].value));
			}
		}
	}).catch(function (err) {
		console.error('balanceSheetData Promise Rejected');
		console.error(err);
	});


	// #################################### Profit And Loss ####################################
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/management-overview?p_p_id=go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=PROFIT_LOSS_STMT_FORWARD');
	await page.evaluate(() => {

		var result = {
			rows: [],
		};

		// Left Side
		var titles = document.querySelectorAll('div.portlet-content > div > div > table > tbody > tr > td:nth-child(1)');
		//console.log(titles);
		var datasets = document.querySelectorAll('div.portlet-content > div > div > table > tbody > tr > td:nth-child(2)');
		var i = 0;
		while (i < datasets.length) {

			var title = (titles[i] && titles[i].innerText) || ' ';
			var val = (datasets[i] && datasets[i].innerText) || ' ';

			result.rows.push({
				key: title,
				value: val.replace(/\./g, '').replace(',', '.')
			});
			i++;
		}
		return Promise.resolve(result);
	}).then(function (result) {
		//console.log(result);
		var sheet = wb.addWorksheet('Profit and Loss');
		for (var i = 0; i < result.rows.length; i++) {
			// Write title to column A
			sheet.cell(1 + i, 1).string(result.rows[i].key);
			// Check if its a number
			// write results in column B+STEP
			if (result.rows[i].value === '' || isNaN(result.rows[i].value)) {
				//console.log(result.rows[i].value);
			} else {
				sheet.cell(1 + i, 1 + STEP).number(Number(result.rows[i].value));
			}
		}
	}).catch(function (err) {
		console.error('sheet Promise Rejected');
		console.error(err);
	});


	// #################################### Process Development ####################################
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/management-overview?p_p_id=go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=PROCESS_DEVELOPMENT_FORWARD');
	await page.evaluate(() => {

		var result = {
			rows: [],
		};

		// Get the table selector
		var carFinancing = document.querySelectorAll('#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > table:nth-child(5)')[0];

		// Get the headline
		var headline = document.querySelectorAll('#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > table:nth-child(5)')[0].children[0].children[0].innerText

		result.rows.push({
			key: headline,
			value: ''
		});

		var childTables = carFinancing.children[0].children[1].children;


		for (var j = 0; j < childTables.length; j++) {
			//console.log(childTables[j])
			// td -> table -> tbody ->
			var tdChilds = childTables[j].children[0].children[0].children;
			//console.log(tdChilds)
			// Start at 2 because of headlines
			result.rows.push({
				key: tdChilds[0].children[0].innerText,
				value: ''
			});
			for (var i = 2; i < tdChilds.length; i++) {
				//console.log(tdChilds[i].children[0].innerText);
				//console.log(tdChilds[i].children[1].innerText);
				var key = tdChilds[i].children[0].innerText;
				var value = tdChilds[i].children[1].innerText;
				// Use regex to match the unit
				var matches = value.match(/([0-9\.,]+) ([\%FTAMLOCIPE\/yr]+)/);

				if (matches[0]) {
					key += ' - ' + matches[2];
				}
				if (matches[1]) {
					value = matches[1];
				}

				result.rows.push({
					key: key,
					value: value.replace(/\./g, '').replace(',', '.'),
				});
			}
		}

		result.rows.push({
			key: '',
			value: ''
		});

		// Get the table selector
		var savings = document.querySelectorAll('#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > table:nth-child(7)')[0];

		// Get the headline
		headline = document.querySelectorAll('#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > table:nth-child(7)')[0].children[0].children[0].innerText

		result.rows.push({
			key: headline,
			value: ''
		});

		childTables = savings.children[0].children[1].children;


		for (j = 0; j < childTables.length; j++) {
			//console.log(childTables[j])
			// td -> table -> tbody ->
			tdChilds = childTables[j].children[0].children[0].children;
			//console.log(tdChilds)
			// Start at 2 because of headlines
			result.rows.push({
				key: tdChilds[0].children[0].innerText,
				value: ''
			});
			for (i = 2; i < tdChilds.length; i++) {
				//console.log(tdChilds[i].children[0].innerText);
				//console.log(tdChilds[i].children[1].innerText);
				key = tdChilds[i].children[0].innerText;
				value = tdChilds[i].children[1].innerText;
				// Use regex to match the unit
				matches = value.match(/([0-9\.,]+) ([\%FTAMLOCIPE\/yr]+)/);

				if (matches[0]) {
					key += ' - ' + matches[2];
				}
				if (matches[1]) {
					value = matches[1];
				}

				result.rows.push({
					key: key,
					value: value.replace(/\./g, '').replace(',', '.'),
				});
			}
		}

		return Promise.resolve(result);
	}).then(function (result) {
		console.log(result);
		var sheetPD = wb.addWorksheet('Process Development');
		for (var i = 0; i < result.rows.length; i++) {
			// Write title to column A
			if (result.rows[i].key) {
				sheetPD.cell(1 + i, 1).string(result.rows[i].key);
			}
			// Write the value to column B+STEP
			if (result.rows[i].value) {
				console.log(result.rows[i].value);
				sheetPD.cell(1 + i, 1 + STEP).number(Number(result.rows[i].value));
			}
		}
	}).catch(function (err) {
		console.log('ProcessDevelopment Promise Rejected');
		console.error(err);
	});

	var CrawlResourceManagement = function (sheetName) {
		return page.evaluate(() => {

			var result = {
				rows: [],
			};

			var rows = document.querySelectorAll('#portlet-wrapper-go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2 > div.portlet-content > div > div > form > table > tbody > tr');

			for (var i = 0; i < rows.length; i++) {
				if (rows[i].children.length > 3) {
					var key = rows[i].children[0].innerText;
					var value = rows[i].children[1].innerText;
					// Use regex to match the unit
					var matches = value.match(/([0-9\.,]+) ([\%$FTAMLOCIPE\/yr]+)/);
					if (matches && matches[0]) {
						key += ' - ' + matches[2];
					}
					if (matches && matches[1]) {
						value = matches[1];
					}
					result.rows.push({
						key: key,
						value: value.replace(/\./g, '').replace(',', '.'),
					});
				} else {
					result.rows.push({
						key: rows[i].innerText,
						value: '',
					});
				}
			}

			return Promise.resolve(result);
		}).then(function (result) {
			console.log(result);
			var sheetPD = wb.addWorksheet(sheetName);
			for (var i = 0; i < result.rows.length; i++) {
				// Write title to column A
				if (result.rows[i].key) {
					sheetPD.cell(1 + i, 1).string(result.rows[i].key);
				}
				// Write the value to column B+STEP
				if (result.rows[i].value && !isNaN(result.rows[i].value)) {
					console.log(result.rows[i].value);
					sheetPD.cell(1 + i, 1 + STEP).number(Number(result.rows[i].value));
				}
			}
		}).catch(function (err) {
			console.log(sheetName + ' Promise Rejected');
			console.error(err);
		});
	};
	var CrawlPerformanceManagement = function (sheetName) {
		return page.evaluate(() => {

			var result = {
				rows: [],
			};

			var rows = document.querySelectorAll('#portlet-wrapper-go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2 > div.portlet-content > div > div > form > table > tbody > tr > td > table > tbody > tr');

			for (var i = 0; i < rows.length; i++) {
				if (rows[i].children.length > 3) {
					var key = rows[i].children[0].innerText + ' - ' + rows[i].children[5].innerText;
					var value = rows[i].children[4].innerText;
					result.rows.push({
						key: key,
						value: value.replace(/\./g, '').replace(',', '.'),
					});
				} else {
					result.rows.push({
						key: rows[i].innerText,
						value: '',
					});
				}
			}

			return Promise.resolve(result);
		}).then(function (result) {
			console.log(result);
			var sheetPD = wb.addWorksheet(sheetName);
			for (var i = 0; i < result.rows.length; i++) {
				// Write title to column A
				if (result.rows[i].key) {
					sheetPD.cell(1 + i, 1).string(result.rows[i].key);
				}
				// Write the value to column B+STEP
				if (result.rows[i].value && !isNaN(result.rows[i].value)) {
					console.log(result.rows[i].value);
					sheetPD.cell(1 + i, 1 + STEP).number(Number(result.rows[i].value));
				}
			}
		}).catch(function (err) {
			console.log(sheetName + ' Promise Rejected');
			console.error(err);
		});
	};

	// #################################### CIO ####################################
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&p_p_col_id=column-2&p_p_col_count=1&CURRENT_ROLE=CIO');
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=PERFORMANCE_MEASUREMENT_FORWARD');
	await CrawlPerformanceManagement('CIO - Performance');
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=RESOURCE_MGMT_FORWARD');
	await CrawlResourceManagement('CIO - Resources');

	// #################################### CFO ####################################
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&CURRENT_ROLE=CFO');
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=PERFORMANCE_MEASUREMENT_FORWARD');
	await CrawlPerformanceManagement('CFO - Performance');
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=RESOURCE_MGMT_FORWARD');
	await CrawlResourceManagement('CFO - Resources');

	// #################################### CMO ####################################
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&CURRENT_ROLE=CMO');
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=PERFORMANCE_MEASUREMENT_FORWARD');
	await CrawlPerformanceManagement('CMO - Performance');
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=RESOURCE_MGMT_FORWARD');
	await CrawlResourceManagement('CMO - Resources');

	// #################################### COO ####################################
	await page.goto('http://tum.go4c.org/group/'+CREDS.group+'/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&CURRENT_ROLE=COO');
	await page.goto('http://tum.go4c.org/group/' + CREDS.group + '/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=PERFORMANCE_MEASUREMENT_FORWARD');
	await CrawlPerformanceManagement('COO - Performance');
	await page.goto('http://tum.go4c.org/group/'+CREDS.group+'/balance-scorecard?p_p_id=go4c_bsc_rm_portlet_WAR_go4cbscrmportlet_INSTANCE_s6L2&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=RESOURCE_MGMT_FORWARD');
	await CrawlResourceManagement('COO - Resources');


	// Add Worksheets to the workbook
	wb.write('' + STEP + '-Periode.xlsx');

	// close browser
	await browser.close();
}

getData();