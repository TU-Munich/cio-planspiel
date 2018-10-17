const puppeteer = require('puppeteer');
const CREDS = require('./creds.json');
// Require library
const request = require('request');
const fs = require('fs');
var async = require('async');

async function cleanNumbers(dirty) {
    var result = dirty.replace('.', '').replace(',', '.');
    return Number(result)
}

async function getPics() {
    const STEP = 10;
    const list = [
        "ratio_total_cost_income",
        "net_income",
        "ratio_it_cost_inc",
        "cash_reserves",
        "t_bills",
        "t_bonds",
        "reverse_repos",
        "bank_loans_assets",
        "car_financing_loans",
        "central_bank_advance",
        "repos",
        //"bank_loans_liabilites",
        "customer_savings",
        "equity",
        "interest_income",
        "interest_expense",
        "net_interest_revenues",
        "loan_loss_provisions",
        "net_interest_revenues_after_loan_loss",
        "administrative_costs_management_staff",
        "administrative_costs_functional_staff",
        "administrative_costs_it_staff",
        "marketing_exp",
        "administrative_costs_it_materials",
        "avail_it",
        "avail_it_supp",
        "it_hw_server",
        "it_hw_storage",
        "it_hw_com",
        "it_sw_systemsw",
        "it_sw_ba_total",
        "it_hw_desktop_total",
        "it_sw_office_total",
        "utilization_it_hw_server",
        "utilization_it_hw_storage",
        "utilization_it_hw_com",
        "request_to_be_processed_loans(1)",
        "request_to_be_processed_loans(2)",
        "request_to_be_processed_loans(3)",
        "requests_success_loans(1)",
        "requests_success_loans(2)",
        "requests_success_loans(3)",
        "request_to_be_processed_sav",
        "requests_success_sav",
        "proc_time_per_contract_s_a_m_loans",
        "proc_time_per_contract_origination_loans",
        "proc_time_per_contract_servicing_loans",
        "proc_time_per_contract_s_a_m_sav",
        "proc_time_per_contract_origination_sav",
        "proc_time_per_contract_servicing_sav",
        "error_rate(5)",
        "error_rate(6)",
        "error_rate(7)",
        "error_rate(8)",
        "error_rate(9)",
        "error_rate(10)",
        "customer_satisfaction",
        "ratio_new_car_fin_volume",
        "ratio_new_sav_acc_volume",
        "marketing_efficency_loans",
        "marketing_efficency_sav",
        "emp(1)",
        "emp(2)",
        "emp(3)",
        "emp(4)",
        "emp(5)",
        "emp(6)",
        "emp(7)",
        "emp(8)",
        "emp(9)",
        "emp(10)",
        "emp_it(1)",
        "emp_it(2)",
        "emp_it(3)",
        "emp_it(4)",
        "emp_it(5)",
        "emp_total",
        "sat_emp(1)",
        "sat_emp(2)",
        "sat_emp(3)",
        "sat_emp(4)",
        "sat_emp(5)",
        "sat_emp(6)",
        "sat_emp(7)",
        "sat_emp(8)",
        "sat_emp(9)",
        "sat_emp(10)",
        "sat_emp_it(1)",
        "sat_emp_it(2)",
        "sat_emp_it(3)",
        "sat_emp_it(4)",
        "sat_emp_it(5)",
        "utilization_emp_marketing",
        "utilization_emp(5)",
        "utilization_emp(6)",
        "utilization_emp(7)",
        "utilization_emp(8)",
        "utilization_emp(9)",
        "utilization_emp(10)",
        "utilization_emp_it(1)",
        "utilization_emp_it(2)",
        "utilization_emp_it(3)",
        "utilization_emp_it(4)",
        "utilization_emp_it(5)",
        "sl_emp(5)",
        "sl_emp(6)",
        "sl_emp(7)",
        "sl_emp(8)",
        "sl_emp(9)",
        "sl_emp(10)",
        "sl_emp_it_dev(1)",
        "sl_emp_it_supp(1)",
        "sl_emp_it_dev(2)",
        "sl_emp_it_supp(2)",
        "sl_emp_it_dev(3)",
        "sl_emp_it_supp(3)",
        "sl_emp_it_dev(4)",
        "sl_emp_it_supp(4)",
        "sl_emp_it_dev(5)",
        "sl_emp_it_supp(5)",
    ];

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


    // START
    await page.goto("http://tum.go4c.org/group/"+CREDS.group+"/management-overview?p_p_id=go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8&p_p_lifecycle=0&p_p_state=maximized&p_p_mode=view&TABS_FORWARD=MGMT_REPORT_FORWARD");

    await async.eachSeries(list, function (name, next) {

        async.series([
                function (callback) {
                    page.evaluate((name) => {
                        console.log("START: " + name);

                        let selectSelector = "#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 > div.portlet-content > div > div > form > select";

                        // Remove all selected entries
                        let elements = document.querySelector(selectSelector).options;
                        for (let i = 0; i < elements.length; i++) {
                            if (elements[i]) {
                                elements[i].selected = false;
                            }
                        }
                        document.querySelector(selectSelector + ' > option[value="' + name + '"]').selected = true;
                        let element = document.querySelector(selectSelector);
                        let event = new Event('change', {bubbles: true});
                        event.simulated = true;
                        element.dispatchEvent(event);
                        return Promise.resolve();
                    }, name).then(function () {
                        return callback();
                    });
                },
                function (callback) {
                    page.waitForNavigation().then(() => {
                        callback();
                    })
                },
                function (callback) {
                    page.evaluate(() => {
                        let elem = document.querySelector("#portlet-wrapper-go4c_mgmtoverview_portlet_WAR_go4cmgmtoverviewportlet_INSTANCE_lBa8 img");
                        if (elem && elem.src) {
                            return Promise.resolve(elem.src);
                        }
                        return Promise.resolve(null);
                    }).then((url) => {
                        if (url) {
                            console.log(url);
                            request(url, {encoding: 'binary'}, function (err, response, body) {
                                fs.writeFile("./perioden/" + STEP +"/ " + STEP + "-" + name + '.png', body, 'binary', function (err) {
                                    if (err) {
                                        console.error(err);
                                    }
                                    console.log('File saved.');
                                    callback();
                                })
                            });
                        }
                    });
                }
            ],
            // optional callback
            function (err, results) {
                next(err);
            });

    }, function (err) {
        if (err) {
            console.error(err);
        }
        console.log('iterating done');
    });
    // close browser
    //await browser.close();

}


getPics();