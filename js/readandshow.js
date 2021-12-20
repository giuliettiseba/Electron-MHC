var path = require('path');
var fs = require('fs');
const {Chart} = require("chart.js");
const feather = require("feather-icons");

var obj_file_path = path.join(__dirname, '/json/obj.json');
var desc_file_path = path.join(__dirname, '/json/descriptions.json');
var tests_file_path = path.join(__dirname, 'json/test.json');

var data;
var tests;
var descriptions
var class_test

$(".nav-link").on("click", function () {
    $(".navbar-nav").find(".active").removeClass("active");
    $(this).addClass("active");
});


$("#click_licences").click(function () {
    build_accordion(data['Licenses'], descriptions)
});

$("#click_products").click(function () {
    build_accordion(data['Products'], descriptions)
});

$("#click_server_info").click(function () {
    build_accordion(data['ServersInformation'], descriptions)
});


$("#click_performance").click(function () {
    build_charts(data['Performance'], descriptions)
});


$("#click_Report").click(function () {
    build_accordion(data['Report'], descriptions)
});

$("#click_logs").click(function () {
    build_accordion(data['ServersInformation'], descriptions)
});

$(document).ready(function () {
    loadFiles()
});

function loadFiles() {
    data = JSON.parse(fs.readFileSync(obj_file_path).toString())
    descriptions = JSON.parse(fs.readFileSync(desc_file_path).toString())
    tests = JSON.parse(fs.readFileSync(tests_file_path).toString())
}

function initChart(chartName, dataKey, dataValue, max) {
    feather.replace({'aria-hidden': 'true'})

    var ctx = document.getElementById(chartName)
    // eslint-disable-next-line no-unused-vars
    const data = {
        labels: dataKey,
        datasets: [{
            //     label: chartName,
            data: dataValue,
            fill: false,
            borderColor: 'rgb(75, 192, 192)',
            tension: 0.1,
            stepped: true,
        }]
    };

    const options = {
        plugins: {
            legend: {
                display: false,
            }
        },
        responsive: true,
        scales: {
            y: {
               // min: 0,
                max: max,

            },
            x: {
                display: false,
            }

        }
    }


    const config = {
        type: 'line',
        data: data,
        options: options
    };

    new Chart(ctx, config)

}


function buildCard(chartName, dataKey, dataValue) {

    const col = document.createElement('div');
    $(col).addClass('col')

    const card = document.createElement('div')
    $(card).addClass('card').addClass('shadow-sm').appendTo(col)

    const canvas = document.createElement('canvas')
    $(canvas).addClass('my-4 w-100')
        .attr({
            id: chartName,
            with: 200,
            height: 100
        })
        .appendTo(card)

    const card_body = document.createElement('div')
    $(card_body).addClass('card-body').appendTo(card)

    const text = document.createElement('p')
    $(text).addClass('card-text').html(chartName).appendTo(card_body)

    return col
}

function build_chart(chartName, dataKey, dataValue, max) {
    $(`#${chartName}`).remove();
    var canvas = document.createElement('canvas');
    canvas.id = chartName;
    canvas.width = 50;
    canvas.height = 50;
    canvas.style.zIndex = 0;
    canvas.style.position = "absolute";
    canvas.style.border = "1px solid";

    const card = buildCard(chartName, dataKey, dataValue)

    $('#graphContainer').append(card);

    initChart(chartName, dataKey, dataValue, max);
}


function cleanContainers() {
    $("#accordion").html('')
    $("#graphContainer").html('')
}

function build_charts(data, descriptions) {
    cleanContainers();

    Object.entries(data).forEach((performanceCounter, index) => {
        const perfName = performanceCounter[0]
        performanceCounter[1].forEach((computer, index) => {
            const chartName = `${perfName}_${computer["PSComputerName"]}`
            const dataValue = computer['Samples'].map(a => a['Value'])
            const dataKey = computer['Samples'].map(a => a['TimeStamp'].replace('Z', ''));
            build_chart(chartName, dataKey, dataValue, computer['Max']);
        })
    })
}


function build_accordion(data, descriptions) {

    cleanContainers();
    $.each(data, function (e, o) {

        const accordion_item = document.createElement('div');
        $(accordion_item).addClass("accordion-item")
            .appendTo($("#accordion")) //main div

        const item_header = document.createElement('h2');
        $(item_header).addClass("accordion-header")
            .attr('id', `heading_${e}`)
            .appendTo(accordion_item)

        const des = e
//      const des = descriptions[e]?.Spanish.Title ? descriptions[e]?.Spanish.Title : e /// DEBUG MODE ON

        const item_button = document.createElement('button');
        $(item_button).addClass("accordion-button")
            .attr({
                'type': `button`,
                'data-bs-toggle': 'collapse',
                'data-bs-target': `#collapse_${e}`,
                'aria-expanded': 'true',
                'aria-controls': `collapse_${e}`
            }).text(des)
            .appendTo(item_header)

        const item_collapse = document.createElement('div');
        $(item_collapse).addClass("accordion-collapse")
            .addClass("collapse")
            .addClass("show")
            .attr({
                'id': `collapse_${e}`,
                'aria-labelledby': `heading_${e}`,
                //'data-bs-parent': '#accordion',
            }).appendTo(accordion_item)

        const item_body = document.createElement('div');
        $(item_body).addClass("accordion-body")
            .appendTo(item_collapse)
            .next(show_description(e, descriptions, item_body))
            .next(show_table(e, o, item_body))
    })
}

function show_description(e, descriptions, item_body) {

    const description_div = document.createElement('div');
    $(description_div).appendTo(item_body)

    const p = document.createElement('p');
    $(p).text(descriptions[e]?.Spanish.Description)
        .appendTo(description_div)

}


function show_table(e, o, item_body) {

    o = Array.isArray(o) ? o : [o];

    const table = document.createElement('table');
    $(table).addClass("table")
        .addClass("table-hover")
        .addClass("table-bordered")
        .attr('id', `id=target_table_${e}`)
        .appendTo(item_body)

    const thead = document.createElement('thead')
    $(thead).appendTo(table)

    const tr = document.createElement('tr')
    $(tr).appendTo(thead)

    /// Headers
    if (o !== undefined && Object.keys(o).length !== 0) {

        $.each(o.at(0), function (k, v) {
            const th = document.createElement('th')
            $(th).text(k)
                .appendTo(tr)
        });
    }

    const trClasses = ["table-primary", "table-secondary", "table-info", "table-light"]
    var dict = {}
    i = 0


    const tbody = document.createElement('tbody')
    $(tbody).appendTo(table)

    class_test = tests[e]

    // Rows
    $.each(o, function () {
        const tr = document.createElement('tr')
        const group_column = (this.PSComputerName != null) ? 'PSComputerName' : 'RecorderName'

        // Group by PSComputerName
        dict[this[group_column]] = dict[this[group_column]] === undefined ? trClasses[i++ % trClasses.length] : dict[this[group_column]]
        $(tr).addClass(dict[this[group_column]])
            .appendTo(tbody)


        $.each(this, function (k, v) {
            const td = document.createElement('td')


            const _td = $(td).text(v)

            if (class_test !== undefined && class_test[k] !== undefined) {
                if (class_test[k] === "isPassDate") {
                    const date = v.split('/')
                    if (date.length === 3) {
                        const exp = new Date(date[2], date[0] - 1, date[1]);
                        const now = new Date();
                        (now > exp) ? _td.addClass('table-danger') : _td.addClass('table-success')
                    }
                } else if (class_test[k] === "isLessThan65536") {
                    (65536 > v) ? _td.addClass('table-danger') : _td.addClass('table-success')
                } else if (class_test[k] === "isLessThan1000") {
                    (1000 > v) ? _td.addClass('table-danger') : _td.addClass('table-success')
                } else {
                    (class_test[k] != String(v)) ? _td.addClass('table-danger') : _td.addClass('table-success')
                }
            }
            _td.appendTo(tr)
        });
    });
}

function getDateIfDate(d) {
    var m = d.match(/\/Date\((\d+)\)\//);
    return m ? (new Date(+m[1])).toLocaleDateString('en-US', {month: '2-digit', day: '2-digit', year: 'numeric'}) : d;
}


