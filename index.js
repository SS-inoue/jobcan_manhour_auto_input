const webdriver = require('selenium-webdriver');
const { Builder, By, until } = webdriver;
const assert = require("assert");
let fs = require('fs');
const path = require('path');
let XLSX = require('xlsx');
let driver;

let workbook = XLSX.readFile('工数入力.xlsx', {cellDates:true})
let xlsx;
let newXlsx = {};

let targetDate;
let MODE_DRYRUN = false;
let MODE_TEST = false;
let account;
let password;

for(var i = 0;i < process.argv.length; i++){
	console.log("argv[" + i + "] = " + process.argv[i]);
	if(process.argv[i] === '--mode') {
		if (process.argv[i + 1] === 'dryrun') {
			MODE_DRYRUN = true;
		}
		if (process.argv[i + 1] === 'test') {
			MODE_TEST = true;
		}
	}
	if(process.argv[i] === '--sheet') {
		targetDate = process.argv[i + 1];
	}
	if(process.argv[i] === '--account') {
		account = process.argv[i + 1];
	}
	if(process.argv[i] === '--password') {
		password = process.argv[i + 1];
	}
}

const match_target = targetDate.match(/([0-9]{4})([0-9]{2})/);

console.log('MODE_DRYRUN', MODE_DRYRUN);

var filepath = `./log/${targetDate}.log`;
 
var dirname = path.dirname(filepath);


fs.access(dirname, fs.constants.R_OK | fs.constants.W_OK, (error) => {
	if (error) {
		if (error.code === "ENOENT") {
			fs.mkdirSync(dirname);
		} else {
			return;
		}
	}
	// fs.writeFile(filepath, "Hello World !", "utf8", (error) => { });
});

function appendLog(txt) {
	const runTime = new Date();
	const message = `[${runTime.getFullYear()}.${runTime.getMonth() + 1}.${runTime.getDate()} ${runTime.getHours()}:${runTime.getMinutes()}] ${txt}\n`;
	fs.appendFile(filepath, message, (err) => {
		if (err) throw err;
		console.log(message);
	});
}

workbook.SheetNames.forEach(sheet => {
    if(targetDate == sheet) xlsx = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
})

const dist = path.join(process.cwd(), 'dist');

fs.writeFile(path.join(dist, 'xlsx.json'), JSON.stringify(xlsx, null, '    '), (err)=>{
	if(err) console.log(`error!::${err}`);
});

for (item of xlsx) {
	let itemDate = new Date(item["start date"]).toLocaleString({ timeZone: 'Asia/Tokyo' });
	let matches_date = itemDate.match(/([0-9]{4})-([0-9]+)-([0-9]+) 0:00:00/)
	let newKey = `${matches_date[1]}${("0"+matches_date[2]).slice(-2)}${("0"+matches_date[3]).slice(-2)}`;
	if (typeof newXlsx[newKey] === "undefined") {
		newXlsx[newKey] = []
	}
	newXlsx[newKey].push(item);
}

fs.writeFile(path.join(dist, 'newxlsx.json'), JSON.stringify(newXlsx, null, '    '), (err)=>{
	if(err) console.log(`error!::${err}`);
});


describe("SeleniumChromeTest", () => {
	before(() => {
		appendLog('処理を開始しました')
		const capabilities = webdriver.Capabilities.chrome();
		capabilities.set('chromeOptions', {
		args: [
			'--disable-headless-mode',
			// '--no-sandbox',
			// '--disable-gpu',
			`--window-size=1980,1200`
			// other chrome options
		]
		});
		driver = new Builder().withCapabilities(capabilities).build();
	});
	after(() => {
		appendLog('処理を終了しました')
	  return driver.quit();
	});
	it("正常系_表示_ページタイトル", async () => {
		// ログイン
		await driver.get("https://id.jobcan.jp/users/sign_in?app_key=atd&redirect_to=https://ssl.jobcan.jp/jbcoauth/callback");

		const title = await driver.getTitle();
		assert.equal(title, "ジョブカン共通IDログイン");

		await driver.findElement(By.name("user[email]")).sendKeys(`${account}\n`);
		await driver.findElement(By.name("user[password]")).sendKeys(`${password}\n`);
		await driver.wait(until.urlIs('https://id.jobcan.jp/account/profile'), 10000);

		// ログイン後のページ1
		const title2 = await driver.getTitle();
		assert.equal(title2, "アカウント情報 | ジョブカン共通ID管理画面");

		await new Promise(resolve => setTimeout(resolve, 1000))
		await driver.findElement(By.className('jbc-app-link')).click();

		// ログイン後のページ2
		const handles = await driver.getAllWindowHandles()
		if (handles.length >= 2) {
			const handle = handles[handles.length - 1];
			await driver.switchTo().window(handle);
		}
		const title3 = await driver.getTitle();
		assert.equal(title3, "JOBCAN MyPage: 井上 恵介");

		let match_item_key;
		let itemData;
		for (key in newXlsx) {
			await driver.findElement(By.id('menu_man_hour_manage_img')).click();
			await driver.findElement(By.css('#menu_man_hour_manage > a:first-child')).click();
			await new Promise(resolve => setTimeout(resolve, 1000))

			match_item_key = key.match(/([0-9]{4})([0-9]{2})([0-9]{2})/)
			itemData = newXlsx[key];

			// 工数管理ページ
			await driver.findElement( By.css(`#search-term select[name="year"] option[value="${match_target[1]}"]`) ).click();

			await new Promise(resolve => setTimeout(resolve, 1000))

			await driver.findElement( By.css(`#search-term select[name="month"] option[value="${Number(match_target[2])}"]`) ).click();
			await new Promise(resolve => setTimeout(resolve, 1000))

			const elements = await driver.findElements( By.css('#search-result table > tbody > tr') );

			let counter = 1;
			for (let element of elements) {
				let date = await driver.findElement( By.css(`#search-result table > tbody > tr:nth-child(${counter}) > td:nth-child(1)`) ).getText();
				let time = await driver.findElement( By.css(`#search-result table > tbody > tr:nth-child(${counter}) > td:nth-child(2)`) ).getText();
				console.log(date, time, date.indexOf(`${match_item_key[2]}/${match_item_key[3]}`));
				if (date.indexOf(`${match_item_key[2]}/${match_item_key[3]}`) >= 0) {
					await driver.findElement( By.css(`#search-result table > tbody > tr:nth-child(${counter}) > td:nth-child(4) .btn`) ).click();
					await driver.wait(until.elementLocated(By.id('save-form')), 10000);

					// 入力モーダル
					let index = 1;
					for (let item of itemData) {
						if (typeof item["skip"] === "undefined" || (typeof item["skip"] !== "undefined" && item["skip"] !== "○")) {
							await driver.findElement(By.css('#edit-menu-contents table tbody > tr:first-child .btn')).click();
							await driver.wait(until.elementLocated(By.css(`tr.daily[data-index="${index}"]`)), 10000);
							let optionsProjects = await driver.findElements(By.css('#edit-menu-contents table tbody > tr:last-child select[name="projects[]"] > option'));
							let counterP = 1;
							for (let option of optionsProjects) {
								let optionText1 = await driver.findElement( By.css(`#edit-menu-contents table tbody > tr:last-child select[name="projects[]"] > option:nth-child(${counterP})`) ).getText();
								if(optionText1 === item["プロジェクト"]) {
									await driver.findElement( By.css(`#edit-menu-contents table tbody > tr:last-child select[name="projects[]"] > option:nth-child(${counterP})`) ).click();
								}
								counterP++;
							}

							await new Promise(resolve => setTimeout(resolve, 1000))
							let optionsTasks = await driver.findElements(By.css('#edit-menu-contents table tbody > tr:last-child select[name="tasks[]"] > option'));
							let counterT = 1;
							for (let option of optionsTasks) {
								let optionText2 = await driver.findElement( By.css(`#edit-menu-contents table tbody > tr:last-child select[name="tasks[]"] > option:nth-child(${counterT})`) ).getText();
								if(optionText2 === item["タスク"]) {
									await driver.findElement( By.css(`#edit-menu-contents table tbody > tr:last-child select[name="tasks[]"] > option:nth-child(${counterT})`) ).click();
								}
								counterT++;
							}
							await new Promise(resolve => setTimeout(resolve, 1000))
							await driver.findElement(By.css('#edit-menu-contents table tbody > tr:last-child input[name="minutes[]"]')).sendKeys(item["時間"]);
							await new Promise(resolve => setTimeout(resolve, 1000))
							await new Promise(resolve => setTimeout(resolve, 500))
							await driver.findElement( By.id(`edit-menu`) ).click();
						}
					}
					
					await new Promise(resolve => setTimeout(resolve, 500))
					let modalTitle = await driver.findElement(By.id('edit-menu-title')).getText();
					console.log(modalTitle);
					let modalTime = await driver.findElement(By.id('un-match-time')).getText();
					console.log(modalTime);
					let sabunTime = modalTime.match(/.+([0-9]+:[0-9]+)/)
					console.log(sabunTime);
					if (sabunTime) {
						await driver.findElement(By.css('#edit-menu-contents table tbody > tr:first-child .btn')).click();
						await driver.wait(until.elementLocated(By.css(`tr.daily[data-index="${index}"]`)), 10000);
						let optionsProjects = await driver.findElements(By.css('#edit-menu-contents table tbody > tr:last-child select[name="projects[]"] > option'));
						let counterP = 1;
						for (let option of optionsProjects) {
							let optionText1 = await driver.findElement( By.css(`#edit-menu-contents table tbody > tr:last-child select[name="projects[]"] > option:nth-child(${counterP})`) ).getText();
							if(optionText1 === '[PJ外] 制作チーム共通業務(2021年度)') {
								await driver.findElement( By.css(`#edit-menu-contents table tbody > tr:last-child select[name="projects[]"] > option:nth-child(${counterP})`) ).click();
							}
							counterP++;
						}
						await new Promise(resolve => setTimeout(resolve, 1000))
						let optionsTasks = await driver.findElements(By.css('#edit-menu-contents table tbody > tr:last-child select[name="tasks[]"] > option'));
						let counterT = 1;
						for (let option of optionsTasks) {
							let optionText2 = await driver.findElement( By.css(`#edit-menu-contents table tbody > tr:last-child select[name="tasks[]"] > option:nth-child(${counterT})`) ).getText();
							if(optionText2 === '[制作チーム] 社内業務/社内会議/雑務/メール/Slack/朝礼') {
								await driver.findElement( By.css(`#edit-menu-contents table tbody > tr:last-child select[name="tasks[]"] > option:nth-child(${counterT})`) ).click();
							}
							counterT++;
						}
					
						await new Promise(resolve => setTimeout(resolve, 1000))
					
						await driver.findElement(By.css('#edit-menu-contents table tbody > tr:last-child input[name="minutes[]"]')).sendKeys(sabunTime[1]);
					}
					await new Promise(resolve => setTimeout(resolve, 500))
					await driver.findElement( By.id(`edit-menu`) ).click();
					await new Promise(resolve => setTimeout(resolve, 1000))

					let modalTime2 = await driver.findElement(By.id('un-match-time')).getText();

					let sabunTime2 = modalTime2.match(/.+([0-9]+:[0-9]+)/)
					console.log('modalTime2', modalTime2);
					console.log('sabunTime2', sabunTime2);
					let sabunTime2Txt;
					if (!sabunTime2) {
						sabunTime2Txt = 'なし'
					} else {
						sabunTime2Txt = sabunTime2[1]
					}
					appendLog(`【書き込み完了】${modalTitle}｜${modalTime}|最終差分（${sabunTime2Txt}）`)

					await new Promise(resolve => setTimeout(resolve, 1000))

					if (MODE_DRYRUN) {
						await driver.findElement(By.id('menu-close')).click();
					} else {
						await driver.findElement(By.id('save')).click();
						await new Promise(resolve => setTimeout(resolve, 1000))
						break;
					}
				}
				counter++;
			}
			await new Promise(resolve => setTimeout(resolve, 3000))
		}
	});
});