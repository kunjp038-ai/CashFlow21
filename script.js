/* =====================================================
   CASH FLOW DASHBOARD
   Part 2A-1
   File Upload + Excel/CSV Parsing
===================================================== */


/* ================= GLOBAL DATA ================= */

let rawData = [];

let filteredData = [];

let workbookData = [];



/* ================= FILE UPLOAD ================= */

const excelFile = document.getElementById("excelFile");


excelFile.addEventListener("change", function(event){

    const file = event.target.files[0];


    if(!file){

        return;

    }


    const reader = new FileReader();


    reader.onload = function(e){


        const data = new Uint8Array(e.target.result);


        const workbook = XLSX.read(data, {

            type:"array"

        });


        // Read first sheet

        const sheetName = workbook.SheetNames[0];


        const worksheet = workbook.Sheets[sheetName];


        const jsonData = XLSX.utils.sheet_to_json(
            worksheet,
            {
                defval:""
            }
        );


        processExcelData(jsonData);


    };


    reader.readAsArrayBuffer(file);


});


const API_URL = "https://script.google.com/macros/s/AKfycbw3DT3vhnQqBU2uJzGXLUySk_6BHlfUsvst9aVp3dA3SZ7eSRYxZB93u5mOvRwb98Ld/exec";

document.getElementById("loadGoogleData")
    .addEventListener("click", loadGoogleData);

async function loadGoogleData() {
    const status = document.getElementById("syncStatus");
    status.innerHTML = "Loading...";

    try {
        const response = await fetch(API_URL);

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();

        console.log("Data received:", data);

        // TEMP: show data on screen
        document.getElementById("output").innerHTML =
            `<pre>${JSON.stringify(data, null, 2)}</pre>`;

        status.innerHTML = "✅ Data Loaded";

    } catch (error) {
        console.error(error);
        status.innerHTML = "❌ Failed: " + error.message;
    }
}



/* ================= PROCESS EXCEL DATA ================= */


function processExcelData(data){


    if(!data || data.length===0){

        alert("No data found in file");

        return;

    }


    rawData = data.map(row => normalizeRow(row));


    filteredData = [...rawData];


    console.log("Loaded Data:", rawData);


    initializeFilters();


    updateDashboard();


}



/* ================= NORMALIZE ROW ================= */


function normalizeRow(row){


    return {


        date:
            formatDate(row.Date),


        particular:
            row.Particular || "",


        entry:
            cleanText(row.Entry),


        mainCategory:
            cleanText(row["Main Category"]),


        category:
            cleanText(row.Category),


        subCategory:
            cleanText(row["Sub Category"]),


        costCenter:
            cleanText(row["Cost Center"]),


        mode:
            cleanText(row["Mode Particular"]),


        amount:
            Number(row.Amount) || 0,


        remarks:
            row.Remarks || "",


        month:
            cleanText(row.Month),


        allocation:
            cleanText(row.Allocation),


        fAmount:
            Number(
                row["F. Amount"]
            ) || 0,


        key:
            row.Key || ""

    };


}




/* ================= TEXT CLEANING ================= */


function cleanText(value){


    if(value===undefined || value===null){

        return "";

    }


    return String(value).trim();


}





/* ================= DATE FORMAT ================= */


function formatDate(value){


    if(!value){

        return "";

    }


    try{


        // Excel date number

        if(typeof value === "number"){


            const date =
                XLSX.SSF.parse_date_code(value);


            return `${String(date.d).padStart(2,"0")}/${
                String(date.m).padStart(2,"0")
            }/${date.y}`;


        }



        let date = new Date(value);



        if(!isNaN(date)){


            return `${String(date.getDate()).padStart(2,"0")}/${
                String(date.getMonth()+1).padStart(2,"0")
            }/${date.getFullYear()}`;


        }


    }

    catch(error){


        console.log(error);


    }


    return value;


}





/* ================= NUMBER FORMAT ================= */


function formatCurrency(value){


    return new Intl.NumberFormat(
        "en-IN",
        {

            style:"currency",

            currency:"INR",

            maximumFractionDigits:0

        }

    ).format(value);



}



/* ================= DATA SUMMARY ================= */


function getExpenseData(){


    return filteredData.filter(row =>

        row.entry.toLowerCase()==="debit"

    );


}



function getIncomeData(){


    return filteredData.filter(row =>

        row.entry.toLowerCase()==="credit"

    );


}



/* ================= BASIC TOTAL FUNCTIONS ================= */


function totalIncome(){


    return getIncomeData()

        .reduce(
            (sum,row)=>sum + Math.abs(row.fAmount),
            0
        );


}



function totalExpense(){


    return getExpenseData()

        .reduce(
            (sum,row)=>sum + Math.abs(row.fAmount),
            0
        );


}



function netCashFlow(){


    return totalIncome() - totalExpense();


}

/* =====================================================
   CASH FLOW DASHBOARD
   Part 2A-2
   Filters + KPI + Table + Chart Data Preparation
===================================================== */


/* ================= FILTER INITIALIZATION ================= */

function initializeFilters(){

    populateFilter(
        "monthFilter",
        rawData.map(row=>row.month)
    );


    populateFilter(
        "entryFilter",
        rawData.map(row=>row.entry)
    );


    populateFilter(
        "categoryFilter",
        rawData.map(row=>row.category)
    );


    populateFilter(
        "mainCategoryFilter",
        rawData.map(row=>row.mainCategory)
    );


    populateFilter(
        "costCenterFilter",
        rawData.map(row=>row.costCenter)
    );


    populateFilter(
        "paymentFilter",
        rawData.map(row=>row.mode)
    );

}



/* ================= CREATE DROPDOWN OPTIONS ================= */


function populateFilter(id, values){


    const select=document.getElementById(id);


    if(!select) return;


    const unique=[

        ...new Set(
            values.filter(v=>v!=="")
        )

    ];


    unique.forEach(value=>{


        let option=document.createElement("option");


        option.value=value;

        option.textContent=value;


        select.appendChild(option);


    });


}



/* ================= FILTER EVENTS ================= */


document.querySelectorAll("select")

.forEach(select=>{


    select.addEventListener(
        "change",
        applyFilters
    );


});




function applyFilters(){


    filteredData = rawData.filter(row=>{


        return (

            (!monthFilter.value ||
            row.month===monthFilter.value)


            &&


            (!entryFilter.value ||
            row.entry===entryFilter.value)


            &&


            (!categoryFilter.value ||
            row.category===categoryFilter.value)


            &&


            (!mainCategoryFilter.value ||
            row.mainCategory===mainCategoryFilter.value)


            &&


            (!costCenterFilter.value ||
            row.costCenter===costCenterFilter.value)


            &&


            (!paymentFilter.value ||
            row.mode===paymentFilter.value)


        );


    });



    updateDashboard();


}



/* ================= ELEMENT REFERENCES ================= */


const monthFilter =
document.getElementById("monthFilter");


const entryFilter =
document.getElementById("entryFilter");


const categoryFilter =
document.getElementById("categoryFilter");


const mainCategoryFilter =
document.getElementById("mainCategoryFilter");


const costCenterFilter =
document.getElementById("costCenterFilter");


const paymentFilter =
document.getElementById("paymentFilter");






/* ================= KPI UPDATE (CONSOLIDATED) ================= */

/*
   NOTE: This replaces the old updateKPIs() + refreshKPICards() split.
   Previously updateDashboard() only called refreshKPICards(), which
   never touched needExpense / wantExpense / transactionCount /
   avgExpense — those cards were stuck at their initial values.
   This single function now updates every KPI card, using the
   "actual" (Internal Transfer excluded) figures for consistency,
   plus the new insight KPI cards.
*/

function updateAllKPIs(){

    const kpi = calculateKPIs();

    const nw = calculateNeedWant();


    document.getElementById(
        "totalIncome"
    ).innerHTML =
        formatCurrency(kpi.income);


    document.getElementById(
        "totalExpense"
    ).innerHTML =
        formatCurrency(kpi.expense);


    document.getElementById(
        "netCash"
    ).innerHTML =
        formatCurrency(kpi.cashFlow);


    document.getElementById(
        "needExpense"
    ).innerHTML =
        formatCurrency(nw.need);


    document.getElementById(
        "wantExpense"
    ).innerHTML =
        formatCurrency(nw.want);


    document.getElementById(
        "savingRate"
    ).innerHTML =
        kpi.savings.toFixed(1)+"%";


    document.getElementById(
        "transactionCount"
    ).innerHTML =
        filteredData.length;


    const actualExpenseCount =
        getActualExpenses().length;

    const avg =
        actualExpenseCount>0 ?
        kpi.expense / actualExpenseCount :
        0;

    document.getElementById(
        "avgExpense"
    ).innerHTML =
        formatCurrency(avg);


    updateInsightKPIs();

}



/* ================= INSIGHT KPI HELPERS ================= */


function getTopCategory(){

    const totals = groupBy("mainCategory");

    const entries = Object.entries(totals);

    if(entries.length===0) return null;

    entries.sort((a,b)=>b[1]-a[1]);

    return { name: entries[0][0], amount: entries[0][1] };

}


function getBiggestExpense(){

    const expenses = getActualExpenses();

    if(expenses.length===0) return null;

    let biggest = expenses[0];

    expenses.forEach(row=>{

        if(Math.abs(row.fAmount) > Math.abs(biggest.fAmount)){

            biggest = row;

        }

    });

    return biggest;

}


function getTopPaymentMode(){

    const totals = groupBy("mode");

    const entries = Object.entries(totals);

    if(entries.length===0) return null;

    entries.sort((a,b)=>b[1]-a[1]);

    return { name: entries[0][0], amount: entries[0][1] };

}


function getHighestSpendingMonth(){

    const totals = {};

    getActualExpenses().forEach(row=>{

        const key = row.month || "Unknown";

        totals[key] = (totals[key]||0) + Math.abs(row.fAmount);

    });

    const entries = Object.entries(totals);

    if(entries.length===0) return null;

    entries.sort((a,b)=>b[1]-a[1]);

    return { name: entries[0][0], amount: entries[0][1] };

}



/* ================= INSIGHT KPI UPDATE ================= */


function updateInsightKPIs(){

    const topCat = getTopCategory();

    document.getElementById("topCategory").innerHTML =
        topCat ? topCat.name : "—";

    document.getElementById("topCategoryAmount").innerHTML =
        topCat ? formatCurrency(topCat.amount) : "";


    const biggest = getBiggestExpense();

    document.getElementById("biggestExpense").innerHTML =
        biggest ? formatCurrency(Math.abs(biggest.fAmount)) : "₹0";

    document.getElementById("biggestExpenseName").innerHTML =
        biggest ? (biggest.particular || biggest.category) : "";


    const topMode = getTopPaymentMode();

    document.getElementById("topPaymentMode").innerHTML =
        topMode ? topMode.name : "—";

    document.getElementById("topPaymentModeAmount").innerHTML =
        topMode ? formatCurrency(topMode.amount) : "";


    const topMonth = getHighestSpendingMonth();

    document.getElementById("highestMonth").innerHTML =
        topMonth ? topMonth.name : "—";

    document.getElementById("highestMonthAmount").innerHTML =
        topMonth ? formatCurrency(topMonth.amount) : "";

}



/* ================= EXPENSE TABLE ================= */


function createExpenseTable(){


    const tbody =
    document.querySelector(
        "#expenseTable tbody"
    );


    tbody.innerHTML="";



    const expenses =

        getExpenseData()

        .sort(
            (a,b)=>
            Math.abs(b.fAmount)-
            Math.abs(a.fAmount)
        )

        .slice(0,10);



    expenses.forEach(row=>{


        let tr =
        document.createElement("tr");



        tr.innerHTML=`

            <td>${row.date}</td>

            <td>${row.particular}</td>

            <td>${row.category}</td>

            <td>${row.mode}</td>

            <td>
                ${formatCurrency(
                    Math.abs(row.fAmount)
                )}
            </td>

        `;


        tbody.appendChild(tr);



    });


}




/* ================= CHART DATA ================= */


let chartData={};



function prepareChartData(){



    chartData={


        category:{},


        mainCategory:{},


        payment:{},


        allocation:{},


        monthly:{},


        costCenter:{},


        subCategory:{}



    };




    filteredData.forEach(row=>{


        let expense =
        Math.abs(row.fAmount);



        if(row.entry==="Debit"){



            chartData.category[row.category]=

            (chartData.category[row.category]||0)
            +expense;




            chartData.mainCategory[row.mainCategory]=

            (chartData.mainCategory[row.mainCategory]||0)
            +expense;




            chartData.payment[row.mode]=

            (chartData.payment[row.mode]||0)
            +expense;




            chartData.allocation[row.allocation]=

            (chartData.allocation[row.allocation]||0)
            +expense;




            let costCenterKey =
            row.costCenter || "Other";

            chartData.costCenter[costCenterKey]=

            (chartData.costCenter[costCenterKey]||0)
            +expense;




            let subCategoryKey =
            row.subCategory || "Other";

            chartData.subCategory[subCategoryKey]=

            (chartData.subCategory[subCategoryKey]||0)
            +expense;


        }




        if(!chartData.monthly[row.month]){


            chartData.monthly[row.month]={

                income:0,

                expense:0

            };


        }



        if(row.entry==="Credit"){

            chartData.monthly[row.month].income +=
            Math.abs(row.fAmount);

        }


        else{

            chartData.monthly[row.month].expense +=
            expense;

        }



    });



}
/* =====================================================
   CASH FLOW DASHBOARD
   Part 2B
   Advanced KPI Calculations + Filter Enhancements
===================================================== */


/* ================= RESET FILTER ================= */

function resetFilters(){

    document.querySelectorAll("select")
    .forEach(select=>{

        select.selectedIndex=0;

    });


    filteredData=[...rawData];


    updateDashboard();

}



/* ================= IGNORE INTERNAL TRANSFER ================= */

function getActualExpenses(){

    return filteredData.filter(row=>{


        return (

            row.entry==="Debit"

            &&

            row.mainCategory!=="Internal Transfer"

        );


    });


}



function getActualIncome(){

    return filteredData.filter(row=>{


        return (

            row.entry==="Credit"

            &&

            row.mainCategory!=="Internal Transfer"

        );


    });


}




/* ================= ADVANCED KPI ================= */


function calculateKPIs(){


    const income =

    getActualIncome()

    .reduce(

        (sum,row)=>

        sum + Math.abs(row.fAmount),

        0

    );



    const expense =

    getActualExpenses()

    .reduce(

        (sum,row)=>

        sum + Math.abs(row.fAmount),

        0

    );



    const cashFlow =
        income-expense;



    const savings =

        income>0 ?

        ((cashFlow/income)*100)

        :

        0;



    return {


        income,

        expense,

        cashFlow,

        savings


    };


}




/* ================= NEED WANT KPI ================= */


function calculateNeedWant(){


    let need=0;

    let want=0;



    getActualExpenses()

    .forEach(row=>{


        if(row.allocation==="Need"){


            need += Math.abs(row.fAmount);


        }



        if(row.allocation==="Want"){


            want += Math.abs(row.fAmount);


        }



    });



    let total =
        need+want;



    return{


        need,

        want,


        needPercent:

        total?

        ((need/total)*100)

        :

        0,


        wantPercent:

        total?

        ((want/total)*100)

        :

        0


    };


}




/* ================= FUEL KPI ================= */


function calculateFuel(){


    let petrol=0;

    let diesel=0;



    getActualExpenses()

    .forEach(row=>{


        if(
            row.subCategory
            .toLowerCase()
            .includes("petrol")
        ){

            petrol+=Math.abs(row.fAmount);

        }



        if(
            row.subCategory
            .toLowerCase()
            .includes("diesel")
        ){

            diesel+=Math.abs(row.fAmount);

        }



    });



    return{


        petrol,

        diesel,

        total:
        petrol+diesel


    };


}




/* ================= CATEGORY SUMMARY ================= */


function groupBy(field){


    let result={};



    getActualExpenses()

    .forEach(row=>{


        let key=row[field] || "Other";


        result[key] =

        (result[key] || 0)

        +

        Math.abs(row.fAmount);



    });



    return result;


}





/* ================= COST CENTER SUMMARY ================= */


function costCenterSummary(){


    return groupBy("costCenter");


}





/* ================= PAYMENT SUMMARY ================= */


function paymentSummary(){


    return groupBy("mode");


}





/* ================= DATE FILTER ================= */


function filterByDate(startDate,endDate){


    filteredData = rawData.filter(row=>{


        let d =
        new Date(row.date);



        return (

            d >= new Date(startDate)

            &&

            d <= new Date(endDate)

        );


    });



    updateDashboard();


}





/* ================= AUTO UPDATE ================= */

/*
   updateAllKPIs() (defined earlier, near the KPI cards) is the single
   source of truth for every KPI card now — it replaces the old
   refreshKPICards() which only updated 4 of the 12 cards.
*/

function updateDashboard(){


    updateAllKPIs();


    createExpenseTable();


    prepareChartData();


    if(typeof renderCharts==="function"){

        renderCharts();

    }


}

/* =====================================================
   CASH FLOW DASHBOARD
   Part 3
   Chart.js Visualizations + Interactions
===================================================== */


let charts = {};



/* ================= CHART DEFAULT SETTINGS ================= */


function chartOptions(){


    return {

        responsive:true,

        maintainAspectRatio:false,


        plugins:{

            legend:{

                position:"bottom",

                labels:{

                    color:
                    getComputedStyle(
                        document.body
                    )
                    .color

                }

            }


        },

        scales:{


            x:{


                ticks:{


                    color:
                    getComputedStyle(
                        document.body
                    )
                    .color


                }


            },


            y:{


                ticks:{


                    color:
                    getComputedStyle(
                        document.body
                    )
                    .color


                }


            }


        }


    };


}





/* ================= DESTROY OLD CHART ================= */


function destroyChart(id){


    if(charts[id]){

        charts[id].destroy();

    }


}





/* ================= CREATE CHART ================= */


function createChart(id,type,data,options={}){


    destroyChart(id);



    let ctx =
    document.getElementById(id)
    .getContext("2d");



    charts[id]=new Chart(

        ctx,

        {

            type:type,

            data:data,

            options:
            {

                ...chartOptions(),

                ...options

            }


        }


    );


}




/* ================= RENDER ALL CHARTS ================= */


function renderCharts(){


    createIncomeExpenseChart();


    createMainCategoryChart();


    createCategoryChart();


    createNeedWantChart();


    createPaymentChart();


    createMonthlyChart();


    createRunningChart();


    createCostCenterChart();


    createSubCategoryChart();


}





/* ================= INCOME VS EXPENSE ================= */


function createIncomeExpenseChart(){


    let months =
    Object.keys(
        chartData.monthly
    );



    createChart(

        "incomeExpenseChart",

        "bar",

        {


            labels:months,


            datasets:[


                {

                    label:"Income",

                    data:
                    months.map(
                        m=>
                        chartData
                        .monthly[m]
                        .income
                    )


                },


                {

                    label:"Expense",

                    data:
                    months.map(
                        m=>
                        chartData
                        .monthly[m]
                        .expense
                    )


                }


            ]


        }


    );


}






/* ================= MAIN CATEGORY ================= */


function createMainCategoryChart(){


    let labels =
    Object.keys(
        chartData.mainCategory
    );



    createChart(

        "mainCategoryChart",

        "doughnut",

        {


            labels,


            datasets:[

                {

                    data:

                    labels.map(
                        x=>
                        chartData
                        .mainCategory[x]
                    )


                }

            ]

        }


    );


}





/* ================= CATEGORY CHART ================= */


function createCategoryChart(){


    let labels =
    Object.keys(
        chartData.category
    );



    createChart(

        "categoryChart",

        "bar",

        {


            labels,


            datasets:[

                {

                    label:
                    "Expense",

                    data:

                    labels.map(
                        x=>
                        chartData
                        .category[x]
                    )


                }

            ]

        }


    );


}





/* ================= NEED WANT ================= */


function createNeedWantChart(){


    let labels =
    Object.keys(
        chartData.allocation
    );



    createChart(

        "needWantChart",

        "doughnut",

        {


            labels,


            datasets:[

                {

                    data:

                    labels.map(
                        x=>
                        chartData
                        .allocation[x]
                    )

                }

            ]

        }


    );


}





/* ================= PAYMENT MODE ================= */


function createPaymentChart(){


    let labels =
    Object.keys(
        chartData.payment
    );



    createChart(

        "paymentChart",

        "pie",

        {


            labels,


            datasets:[


                {

                    data:

                    labels.map(

                        x=>

                        chartData
                        .payment[x]

                    )


                }


            ]

        }


    );


}





/* ================= MONTHLY CASH FLOW ================= */


function createMonthlyChart(){


    let months =
    Object.keys(
        chartData.monthly
    );



    createChart(

        "monthlyChart",

        "line",

        {


            labels:months,


            datasets:[


                {

                    label:
                    "Cash Flow",


                    data:

                    months.map(

                        m=>

                        chartData
                        .monthly[m]
                        .income

                        -

                        chartData
                        .monthly[m]
                        .expense


                    ),


                    tension:.4


                }


            ]

        }


    );


}





/* ================= RUNNING BALANCE ================= */


function createRunningChart(){


    let months =
    Object.keys(
        chartData.monthly
    );



    let balance=0;


    let running=[];



    months.forEach(month=>{


        balance +=

        chartData.monthly[month].income

        -

        chartData.monthly[month].expense;



        running.push(balance);


    });




    createChart(

        "runningChart",

        "line",

        {


            labels:months,


            datasets:[

                {

                    label:
                    "Balance",

                    data:running,


                    fill:true,


                    tension:.4


                }

            ]

        }


    );


}





/* ================= COST CENTER CHART ================= */


function createCostCenterChart(){


    let labels =
    Object.keys(
        chartData.costCenter
    );


    createChart(

        "costCenterChart",

        "bar",

        {


            labels,


            datasets:[

                {

                    label:
                    "Expense",

                    data:

                    labels.map(
                        x=>
                        chartData
                        .costCenter[x]
                    )


                }

            ]

        }


    );


}




/* ================= SUB-CATEGORY CHART ================= */


function createSubCategoryChart(){


    let entries =
    Object.entries(
        chartData.subCategory
    )
    .sort((a,b)=>b[1]-a[1])
    .slice(0,10);


    let labels =
    entries.map(e=>e[0]);


    let values =
    entries.map(e=>e[1]);


    createChart(

        "subCategoryChart",

        "bar",

        {


            labels,


            datasets:[

                {

                    label:
                    "Expense",

                    data:values


                }

            ]

        },

        {

            indexAxis:"y"

        }


    );


}




/* ================= DARK MODE CHART REFRESH ================= */


const themeBtn =
document.getElementById("themeBtn");



if(themeBtn){


    themeBtn.addEventListener(
        "click",
        ()=>{


            document.body
            .classList
            .toggle("dark");



            renderCharts();


        }

    );


}




/* ================= INITIAL LOAD ================= */


window.addEventListener(
    "load",
    ()=>{


        if(rawData.length>0){

            renderCharts();

        }


    }

);
