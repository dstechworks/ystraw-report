const { Pool } = require('pg');
const XLSX = require('xlsx');
const nodemailer = require("nodemailer");


const transporter = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        user: 'reports@techworks.co.in',
        pass: 'vucto0-socWiz-cifjaj'
    }
});
const pool = new Pool({
    user: "postgres",
    host: 'db.mgampbhmlnalxohuobpr.supabase.co',
    database: "postgres",
    password: 'gplVhDuxLDMeBKxs',
    port: 5432,
});

let today = new Date();
let oneDayMilliseconds = 24 * 60 * 60 * 1000;
let oneDayBefore = new Date(today.getTime() - oneDayMilliseconds);
let year = oneDayBefore.getFullYear();
let month = String(oneDayBefore.getMonth() + 1).padStart(2, "0");
let day = String(oneDayBefore.getDate()).padStart(2, "0");
let firstDate = `${year}-${month}-01`;
let current_date = `${year}-${month}-${day}`;
let dateDifference = new Date(current_date).getDate() + 1 - new Date(firstDate).getDate();

const localWorkbook = XLSX.readFile('/root/YSTRAW-REPORT/YSTRAW_BASE_SHEET.xlsx');
const sheetName = localWorkbook.SheetNames[0];
const worksheet = localWorkbook.Sheets[sheetName];
const workbook = XLSX.utils.book_new();

let response1;
let response2;

async function querydb() {
    response1 = await pool.query(`select * from ystraw_data_table  where custom_date = '${current_date}'`);
    response2 = await pool.query(`select * from ystraw_data_table where custom_date >= '${firstDate}' and custom_date <= '${current_date}'`);
}

async function dailyReportDetailed() {
    let city = '';
    let outletName = '';
    let outletAddress = '';
    let lastAccessed = '';
    const dataArray = [[
        'Display ID',
        'Display Name',
        'Date',
        'Outlet Name',
        'Outlet Address',
        'City',
        'Last Accessed',
        'Runtime'
    ]];
    let dim = [];

    if (response1?.rows?.length > 0) {
        const excelArrayData = XLSX.utils.sheet_to_json(worksheet);

        excelArrayData.forEach((element, index) => {
            let idDataFromDatabase = response1.rows.find(d => d.display_id == element['Display ID']);

            if (idDataFromDatabase) {
                city = element['City'] ? element['City'] : '';
                outletName = element['Outlet Name'] ? element['Outlet Name'] : '';
                outletAddress = element['Outlet Address'] ? element['Outlet Address'] : '';
                lastAccessed = idDataFromDatabase?.last_accessed ? idDataFromDatabase?.last_accessed : '';

                dim.push(idDataFromDatabase?.display_id);
                dim.push(idDataFromDatabase?.display_name);
                dim.push((idDataFromDatabase?.custom_date)?.toString().substring(4, 15));
                dim.push(outletName);
                dim.push(outletAddress);
                dim.push(city);
                dim.push(lastAccessed);
                dim.push(((idDataFromDatabase?.display_count * 15) / 60).toFixed(2));
                dataArray.push(dim);
            }

            // reset values
            city = '';
            outletName = '';
            outletAddress = '';
            lastAccessed = '';
            dim = [];
        });


        const array_to_sheet = XLSX.utils.aoa_to_sheet(dataArray);
        XLSX.utils.book_append_sheet(workbook, array_to_sheet, 'Daily Detailed Report');
        mtdReportDetailed();
    } else {
        console.log('NO DATA GET FROM SERVER !!');
    }
}

const mtdReportDetailed = () => {
    // All variables
    let percentageActiveDays = 0;
    let calculatedDateDiff = 0;
    let metricEfficiency = 0;
    let activeDays = 0;
    let timescore = 0;
    let isactive = 0;
    let bucket = '';
    let remarks = '';
    let percentageOfActiveDays = '';
    let showMetricEfficiency = '';
    let showRunDayScoring = '';
    let showTimeScore = '';
    let totalRunTime = '';
    let averageRunTime = '';
    let rundayscoring = 0;
    let city = '';
    let outletName = '';
    let outletAddress = '';
    let t1 = 0;
    let t2 = 0;
    let cummalativescore;
    let cummalativerating;

    const dataArray = [[
        'Display Name',
        // 'Date', 'City',
        'Outlet Name',
        // 'Outlet Address',
        'Operation Days',
        'Active Days',
        // '% Active Days',
        'Runtime In Hrs',
        // 'Average Daily Runtime',
        // 'New Metric Efficiency',
        // 'Bucket',
        // 'Remarks',
        // 'Active',
        // 'Run Day Scoring',
        // 'Time score',
        // 'Run days Scoring (Max 10) No of days Active/ No of Total days in Month *10 Max Score',
        // 'Run Time Scoring (Max 10) Avg. No of Hour Active/ Avg. 8 Hours Run *10 Max Score',
        // 'Cummalative Score',
        // 'Cummalative Rating'
    ]];
    let dim = [];

    if (response2?.rows?.length > 0) {
        const excelArrayData = XLSX.utils.sheet_to_json(worksheet);

        excelArrayData.forEach(element => {
            const idDataFromDatabase = response2.rows.filter(d => d.display_id == element['Display ID']);

            // console.log(idDataFromDatabase);

            if (idDataFromDatabase === undefined) {
                city = '';
                outletName = '';
                outletAddress = '';
            } else {
                city = element['City'];
                outletName = element['Outlet Name'];
                outletAddress = element['Outlet Address'];
            }

            if (Number(idDataFromDatabase.length) <= 1) {
                calculatedDateDiff = 1;
            } else {
                calculatedDateDiff = dateDifference;
            }

            // console.log(idDataFromDatabase.length);

            if (idDataFromDatabase.length) {
                // Calculate daily run times for each matching database entry
                const dailyRunTimes = idDataFromDatabase.map(entry => {
                    const displayCount = parseInt(entry.display_count);
                    // console.log(entry?.display_count);
                    if (Number(entry?.display_count) > 0) {
                        activeDays++;
                        isactive = 1;
                    }
                    return ((displayCount * 15) / 60).toFixed(2); // Calculate run time in hours
                });

                // Calculate total and average run time for the current Display ID
                totalRunTime = dailyRunTimes.reduce((acc, time) => acc + parseFloat(time), 0).toFixed(2);
                averageRunTime = (totalRunTime / dailyRunTimes.length).toFixed(2);
                percentageOfActiveDays = ((activeDays / calculatedDateDiff) * 100).toFixed(2) + "%";
                percentageActiveDays = ((activeDays / calculatedDateDiff) * 100).toFixed(2);
                metricEfficiency = (activeDays / calculatedDateDiff) * (averageRunTime / 8) * 100;
                showMetricEfficiency = metricEfficiency.toFixed(2) + "%";
                rundayscoring = (activeDays / calculatedDateDiff);
                showRunDayScoring = (rundayscoring * 100).toFixed(2) + '%';
                timescore = averageRunTime / 8;
                showTimeScore = (timescore * 100).toFixed(2) + '%';

                //metricEfficiency
                if (Math.round(metricEfficiency) >= 80) {
                    bucket = "Above 80";
                } if (Math.round(metricEfficiency) < 80 && Math.round(metricEfficiency) >= 50) {
                    bucket = "Below 80";
                } if (Math.round(metricEfficiency) < 50 && Math.round(metricEfficiency) > 0) {
                    bucket = "Below 50";
                } if (Math.round(metricEfficiency) === 0) {
                    bucket = "Zero";
                }

                //remarks 
                if (Math.round(percentageActiveDays) >= 80 && Math.round(averageRunTime) >= 8) {
                    remarks = 'Outstanding';
                } if ((Math.round(percentageActiveDays) >= 70 && Math.round(percentageActiveDays) < 80) && Math.round(averageRunTime) >= 8) {
                    remarks = 'Good';
                } if (Math.round(percentageActiveDays) <= 70 && Math.round(averageRunTime) > 0) {
                    remarks = 'Average';
                } if (Math.round(percentageActiveDays) > 70 && Math.round(averageRunTime) < 8) {
                    remarks = 'Satisfactory';
                } if (Math.round(percentageActiveDays) <= 70 && Math.round(averageRunTime) === 0) {
                    remarks = 'Poor';
                }

                t1 = (activeDays / calculatedDateDiff) * 10;
                if (Math.floor(averageRunTime / 8 * 10) > 10) {
                    t2 = 10;
                } else {
                    t2 = Math.floor(averageRunTime / 8 * 10);
                }

                cummalativescore = Math.floor(t1 + t2);

                if (cummalativescore >= 17 && cummalativescore <= 20) {
                    cummalativerating = 'Very Good'
                } else if (cummalativescore >= 12 && cummalativescore <= 16) {
                    cummalativerating = 'Good'
                } else if (cummalativescore >= 7 && cummalativescore <= 11) {
                    cummalativerating = 'Poor'
                } else if (cummalativescore < 7) {
                    cummalativerating = 'Critical'
                }

                // logs
                // console.log('\n');
                // console.log(`Display ID: ${element['Display ID']}`);
                // console.log(`Operation Days : ${calculatedDateDiff}`);
                // console.log(`Active Days : ${activeDays}`);
                // console.log(`% Active Days : ${percentageOfActiveDays}`);
                // console.log(`Run Time : ${totalRunTime}`);
                // console.log(`Average Daily Run Time: ${averageRunTime} hours`);
                // console.log(`Metric Efficiency : ${showMetricEfficiency}`);
                // console.log(`Bucket : ${bucket}`);
                // console.log(`Remarks : ${remarks}`);
                // console.log(`isActive : ${isactive}`);
                // console.log(`Run Day Scoring : ${showRunDayScoring}`);
                // console.log(`Time Score : ${showTimeScore}`);
                // console.log(`Run days Scoring 1 : ${t1.toFixed(2)}`);
                // console.log(`Run days Scoring 2 : ${t2}`);
                // console.log(`Cummalative Score : ${cummalativescore}`);
                // console.log(`Cummalative Rating : ${cummalativerating}`);


                dim.push(idDataFromDatabase[0]?.display_name);
                // dim.push(current_date);
                // dim.push(city);
                dim.push(outletName);
                // dim.push(outletAddress);
                dim.push(calculatedDateDiff); // Operation Days
                dim.push(activeDays); // Active Days
                // dim.push(percentageOfActiveDays); // % Active Days
                dim.push(totalRunTime); // Runtime
                // dim.push(averageRunTime); // Average Daily Runtime
                // dim.push(showMetricEfficiency); // New Metric Efficiency
                // dim.push(bucket); // Bucket
                // dim.push(remarks); // remarks
                // dim.push(isactive); // Active
                // dim.push(showRunDayScoring); // Run Day Scoring
                // dim.push(showTimeScore); // Time score
                // dim.push(t1.toFixed(2)); // Run days Scoring 1
                // dim.push(t2); // Run days Scoring 2
                // dim.push(cummalativescore); // Cummalative Score
                // dim.push(cummalativerating); // Cummalative Rating

                dataArray.push(dim)


                //Resets all values
                percentageActiveDays = 0;
                calculatedDateDiff = 0;
                metricEfficiency = 0;
                activeDays = 0;
                timescore = 0;
                isactive = 0;
                bucket = '';
                remarks = '';
                percentageOfActiveDays = '';
                showMetricEfficiency = '';
                showRunDayScoring = '';
                showTimeScore = '';
                totalRunTime = '';
                averageRunTime = '';
                rundayscoring = 0;
                city = '';
                outletName = '';
                outletAddress = ''
                t1 = 0;
                t2 = 0;
                cummalativescore;
                cummalativerating;
                dim = [];
            }
        });

        const array_to_sheet = XLSX.utils.aoa_to_sheet(dataArray);
        XLSX.utils.book_append_sheet(workbook, array_to_sheet, 'MTD Detailed Report');
        console.log("REPORT SAVED !!");

    } else {
        console.log('NO DATA GET FROM SERVER !!');
    }
}

async function reportDelivery(params) {
    try {
        // send mail with defined transport object
        const info = await transporter.sendMail({
            from: 'reports@techworks.co.in',
            // to: 'hitesh.kumar@techworks.co.in',
            to: 'vivekmry1995@gmail.com, radhika@ystraw.com, dhruv@techworks.co.in',
            cc: "Aaditya@techworks.co.in, kunal.m@techworks.co.in, rahul.rajput@techworks.co.in, hitesh.kumar@techworks.co.in",
            subject: "YSTRAW REPORT TILL" + current_date, // Subject line
            html: `<h6>Please find the attachment.</h6>
            <p>&nbsp;</p>
            <table style="width: 420px; font-size: 10pt; font-family: Verdana, sans-serif; background: transparent !important;"
                    border="0" cellspacing="0" cellpadding="0">
                    <tbody>
                            <tr>
                                    <td style="width: 160px; font-size: 10pt; font-family: Verdana, sans-serif; vertical-align: top;"
                                            valign="top">
                                            <p style="margin-bottom: 18px; padding: 0px;"><span
                                                            style="font-size: 12pt; font-family: Verdana, sans-serif; color: #183884; font-weight: bold;"><a
                                                                    href="http://contact.techworksworld.com/" target="_"><img
                                                                            style="width: 120px; height: auto; border: 0;"
                                                                            src="https://i.imgur.com/Eie8F53.png" width="120"
                                                                            border="0" /></a></p>
                                            <p
                                                    style="margin-bottom: 0px; padding: 0px; font-family: Verdana, sans-serif; font-size: 9pt; line-height: 12pt;">
                                                    <a style="color: #e25422; text-decoration: none; font-weight: bold;"
                                                            href="http://www.techworksworld.com" target="_"><span
                                                                    style="text-decoration: none; font-size: 9pt; line-height: 12pt; color: #e25422; font-family: Verdana, sans-serif; font-weight: bold;">www.techworksworld.com</span></a>
                                            </p>
                                    </td>
                                    <td style="width: 30px; min-width: 30px; border-right: 1px solid #e25422;">&nbsp;</td>
                                    <td style="width: 30px; min-width: 30px;">&nbsp;</td>
                                    <td style="width: 200px; font-size: 10pt; color: #444444; font-family: Verdana, sans-serif; vertical-align: top;"
                                            valign="top">
                                            <p
                                                    style="font-family: Verdana, sans-serif; padding: 0px; font-size: 9pt; line-height: 14pt; margin-bottom: 14px;">
                                                    <span
                                                            style="font-family: Verdana, sans-serif; font-size: 9pt; line-height: 14pt;"><span
                                                                    style="font-size: 9pt; line-height: 13pt; color: #262626;"><strong>E:
                                                                    </strong></span><a
                                                                    style="font-size: 9pt; color: #262626; text-decoration: none;"
                                                                    href="mailto:support@techworks.co.in"><span
                                                                            style="text-decoration: none; font-size: 9pt; line-height: 14pt; color: #262626; font-family: Verdana, sans-serif;">support@techworks.co.in</span></a><span><br /></span></span><span><span
                                                                    style="font-size: 9pt; color: #262626;"><strong>T:</strong></span><span
                                                                    style="font-size: 9pt; color: #262626;"> (+91) 11 35007205</span><span><br /></span></span><span><span
                                                                    style="font-size: 9pt; color: #262626;"><strong>A:</strong></span><span
                                                                    style="font-size: 9pt; color: #262626;"> O-7,Second Floor Lajpat
                                                                    Nagar-II,<span>,</span></span><span style="color: #262626;">New
                                                                    Delhi-110024,India</span></span></p>
                                            <p style="margin-bottom: 0px; padding: 0px;"><span><a
                                                                    href="https://www.facebook.com/TechworksSolutionsPvtLtd/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/fb.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.linkedin.com/company/ds-techworks-solutions-pvt-ltd/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/ln.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://twitter.com/techworks14" rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/tt.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.youtube.com/@TechworksDigitalSolutions"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/yt.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.instagram.com/techworks140/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/it.png"
                                                                            width="22" border="0" /></a></span></p>
                                    </td>
                            </tr>
                            <tr style="width: 420px;">
                                    <td style="padding-top: 14px;" colspan="4"><a href="https://techworksworld.com/" target="_"><img
                                                            style="width: 420px; height: auto; border: 0;"
                                                            src="https://i.imgur.com/QoPxSPy.png" width="420" border="0" /></a></td>
                            </tr>
                            <tr>
                                    <td style="padding-top: 14px; text-align: justify;" colspan="4">
                                            <table style="width: 420px; font-size: 10pt; font-family: Verdana, sans-serif; background: transparent !important;"
                                                    border="0" cellspacing="0" cellpadding="0">
                                                    <tbody>
                                                            <tr>
                                                                    <td
                                                                            style="font-size: 8pt; color: #b2b2b2; line-height: 9pt; text-align: justify;">
                                                                            The content of this email is confidential and intended
                                                                            for the recipient specified in message only. It is
                                                                            strictly forbidden to share any part of this message
                                                                            with any third party,without a written consent of the
                                                                            sender. If you received this message by mistake,please
                                                                            reply to this message and follow with its deletion,so
                                                                            that we can ensure such a mistake does not occur in the
                                                                            future.</td>
                                                            </tr>
                                                    </tbody>
                                            </table>
                                    </td>
                            </tr>
                    </tbody>
            </table>`,
            attachments: [
                {   // file on disk as an attachment
                    filename: `YSTRAW REPORT TILL${current_date}.xlsx`,
                    path: `./YSTRAW REPORT TILL${current_date}.xlsx` // stream this file
                }
            ]
        });

        console.log(`Mail Send Succesfull..  (${current_date})`)

    } catch (error) {
        console.log(error);
    }
}

Promise.all([querydb()]).then(() => {
    Promise.all([
        dailyReportDetailed()
    ])
        .then(() => {
            setTimeout(() => {
                const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
                const fileName = 'YSTRAW REPORT TILL' + current_date + '.xlsx'
                XLSX.writeFile(workbook, fileName);
            }, 3000);

            setTimeout(() => {
                reportDelivery();
            }, 10000);
        })
        .catch((error) => {
            console.error("An error occurred:", error);
        })
})