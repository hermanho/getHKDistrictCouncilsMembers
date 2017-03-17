const jsdom = require("jsdom").jsdom;
const fetch = require('node-fetch');
const sleep = require('sleep');
const XLSX = require('xlsx');
const path = require('path');
const ProgressBar = require('progress');

var stringCleanUp = text => {
    return text.replace(/[\t\r\n]/g, '').trim();
}
var emailDecode = a => {
    for (e = '', r = '0x' + a.substr(0, 2) | 0, n = 2; a.length - n; n += 2) {
        e += '%' + ('0' + ('0x' + a.substr(n, 2) ^ r).toString(16)).slice(-2);
    }
    return decodeURIComponent(e);
}
var exportExcel = (data, filename) => {
    var ws_name = "Sheet1";

    /* require XLSX */
    var XLSX = require('xlsx');

    /* set up workbook objects -- some of these will not be required in the future */
    var wb = {};
    wb.Sheets = {};
    wb.Props = {};
    wb.SSF = {};
    wb.SheetNames = [];

    /* create worksheet: */
    var ws = {}

    /* the range object is used to keep track of the range of the sheet */
    var range = {
        s: {
            c: 0,
            r: 0
        },
        e: {
            c: 0,
            r: 0
        }
    };
    if (data.length > 0 && !Array.isArray(data[0])) {
        var keys = Object.keys(data[0]);
        var dummy = {};
        for (var C = 0; C != keys.length; ++C) {
            dummy[keys[C]] = keys[C];
        }
        data.splice(0, 0, dummy);
    }

    /* Iterate through each element in the structure */
    for (var R = 0; R != data.length; ++R) {
        if (range.e.r < R) range.e.r = R;

        var keys = data[R];
        if (!Array.isArray(data[R])) {
            keys = Object.keys(data[R]);
        }
        for (var C = 0; C != keys.length; ++C) {
            if (range.e.c < C) range.e.c = C;

            /* create cell object: .v is the actual data */
            var cell = {
                v: data[R][C]
            };
            if (!Array.isArray(data[R])) {
                cell.v = data[R][keys[C]];
            }
            if (cell.v == null) continue;

            /* create the correct cell reference */
            var cell_ref = XLSX.utils.encode_cell({
                c: C,
                r: R
            });

            /* determine the cell type */
            if (typeof cell.v === 'number') cell.t = 'n';
            else if (typeof cell.v === 'boolean') cell.t = 'b';
            else cell.t = 's';

            /* add to structure */
            ws[cell_ref] = cell;
        }
    }
    ws['!ref'] = XLSX.utils.encode_range(range);

    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;

    /* write file */
    XLSX.writeFile(wb, filename);
}

var domPromise = new Promise((resolve, reject) => {
    jsdom.env('http://www.districtcouncils.gov.hk/index.html',
        function (err, window) {
            const $ = require('jquery')(window);

            const districts = $('[id^="link"]').toArray().map((e) => {
                let href = $(e).attr('href');
                return href.substr(0, href.indexOf('/'));
            });
            resolve(districts);
        }
    );
});

const nameTrimWording = ['議員', '先生', '女士', ',', 'MH', 'JP', 'BBS'];

var processMembers = function (memberUrls) {
    var bar2 = new ProgressBar('Loading member data :current/:total', {
        total: memberUrls.length
    });
    var result = [];
    return memberUrls.reduce(function (p, m) {
        return p.then(function () {
            var p2 = new Promise((resolve, reject) => {
                jsdom.env('http://www.districtcouncils.gov.hk/' + m.district + '/tc_chi/members/info/' + m.memberUrl,
                    function (err, window) {
                        const $ = require('jquery')(window);

                        $.expr[":"].innertext = $.expr.createPseudo(function (arg) {
                            return function (elem) {
                                return $(elem).text().replace(/[\t\r\n]/g, '').trim() == arg;
                            };
                        });
                        var member = {};
                        member.name = stringCleanUp($('.member_name').text());
                        nameTrimWording.forEach((w) => {
                            if (member.name.indexOf(w) > 0) {
                                member.name = member.name.substr(0, member.name.indexOf(w));
                            }
                        });
                        member.district = $('.mySection').text();
                        member.district = member.district.substr(0, member.district.indexOf('區議會'));
                        member.seat = stringCleanUp($(':innertext(席位)').parent().text()).replace('席位', '');
                        member.subDistrict = stringCleanUp($(':innertext(選區)').parent().text()).replace('選區', '');
                        member.occupation = stringCleanUp($(':innertext(職業)').parent().text()).replace('職業', '');
                        member.politicalAffiliation = stringCleanUp($(':innertext(所屬政治聯繫)').parent().text()).replace('所屬政治聯繫', '');
                        member.address = stringCleanUp($(':innertext(地址)').parent().text()).replace('地址', '');
                        member.tel = stringCleanUp($(':innertext(電話)').parent().text()).replace('電話', '');
                        member.url = stringCleanUp($(':innertext(網頁)').parent().text()).replace('網頁', '');
                        var emailsDom = $(':innertext(電郵地址)').parent().find('[data-cfemail]');
                        var emails = []
                        emailsDom.each((i, val) => {
                            emails.push(emailDecode($(val).attr('data-cfemail')));
                        });
                        member.email = emails.join(' / ');
                        result.push(member);
                        bar2.tick();
                        resolve();
                    }
                );
            });
            return p2;
        });
    }, Promise.resolve()).then(_ => {
        return result;
    });
};

var promiseChain = domPromise.then((districts) => {
        var districtsPromises = [];
        var bar1 = new ProgressBar('Loading districts :current/:total', {
            total: districts.length
        });
        for (let d in districts) {
            var p = new Promise((resolve, reject) => {

                jsdom.env('http://www.districtcouncils.gov.hk/' + districts[d] + '/tc_chi/members/info/dc_member_list.php',
                    function (err, window) {
                        const $ = require('jquery')(window);
                        const memberUrls = $('[href^="dc_member_list_detail"]').toArray().map((e) => {
                            return {
                                district: districts[d],
                                memberUrl: $(e).attr('href')
                            };
                        });
                        bar1.tick();
                        resolve(memberUrls);
                    }
                );
            });
            districtsPromises.push(p);
        }
        return Promise.all(districtsPromises);
    })
    .then(memberUrls => {
        return memberUrls.reduce((array, val) => {
            return array.concat(val);
        }, []);
    })
    .then(memberUrls => processMembers(memberUrls))
    .then(members => {
        exportExcel(members, path.resolve('123.xlsx'));
        console.log('done');
        return;
    });

return promiseChain;