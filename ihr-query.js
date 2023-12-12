function getHeader() {
    return {
        "accept": "application/json, text/plain, */*",
        "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "authorization": "Bearer " + sessionStorage.UEditorToken,
        "cache-control": "no-cache",
        "content-type": "application/json;charset=UTF-8",
        "isajax": "true",
        "pragma": "no-cache"
    }
}

async function queryUnitodoHistoricTodo(count) {

    let resp = await fetch("http://ihr.hq.cmcc/cc/hrp-rely/web/task/queryUnitodoHistoricTodo", {
        "headers": getHeader(),
        "referrer": "http://ihr.hq.cmcc/cs/hrp/launchApplication/myTasked",
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": JSON.stringify({
            "processInstName": "考核打分汇总",
            "creatorUname": "",
            "creatorUnumber": "",
            "currentState": "",
            "currentPage": 1,
            "pageSize": count
        }),
        "method": "POST",
        "mode": "cors",
        "credentials": "include"
    });

    let json = await resp.json();

    console.log(json);

    return json.data.list.map((v,i)=>{
        let obj = {
            name: v.processInstName.replace(/月度考核.*/, ''),
            workItemId: v.workItemId,
            processInstID: v.pcUrl.replace(/.*?&processInstID=(\d+)&.*/, "$1")
        }
        return obj;
    }
    );
}

async function queryList(workItemId) {

    let resp = await fetch("http://ihr.hq.cmcc/cc/pms/api/evaluateController/queryList?flowInstanceId=" + workItemId, {
        "headers": getHeader(),
        "referrer": "http://ihr.hq.cmcc/cs/pms/leadership-rating",
        "referrerPolicy": "strict-origin-when-cross-origin",
        "body": "{}",
        "method": "POST",
        "mode": "cors",
        "credentials": "include"
    });

    let json = await resp.json();
    console.log(json);
    let personList = [];
    for (let i = 0; i < json.data.deptGroupDtos.length; i++) {
        for (let j = 0; j < json.data.deptGroupDtos[i].groupEmplDtoList.length; j++) {
            let dept = json.data.deptGroupDtos[i];
            let group = json.data.deptGroupDtos[i].groupEmplDtoList[j];
            group.emplApprDtoList.forEach(empl=>{

                let person = {
                    pmsGroupId: group.pmsGroupId,
                    pmsGroupName: group.pmsGroupName,

                    cmDeptId: empl.cmDeptId,
                    cmDeptName: empl.cmDeptName,
                    cmEmplId: empl.cmEmplId,
                    cmEmplName: empl.cmEmplName,
                    pmsReviewerId: empl.pmsReviewerId,
                    pmsReviewerName: empl.pmsReviewerName,
                    preReviewerId: empl.preReviewerId,
                    preReviewerName: empl.preReviewerName,
                    pmsRate: empl.pmsRate,
                    pmsScore: empl.pmsScore,
                    pmsScorePre: empl.pmsScorePre,
                    pmsScoreSelf: empl.pmsScoreSelf

                }
                personList.push(person);
            }
            );
        }

    }

    return personList;
}

async function queryExcelData(monthCount) {

    let monthArray = await queryUnitodoHistoricTodo(monthCount)
    console.log(monthArray);

    let dataSheet = {};

    for (var i = 0; i < monthArray.length; i++) {
        let v = monthArray[i];
        let month = v.name;
        console.log(v.processInstID);
        let personList = await queryList(v.processInstID);

        console.log(personList)
        personList.forEach(v=>{

            let data = dataSheet[v.cmEmplId];
            if (!data) {
                data = dataSheet[v.cmEmplId] = v;
            }

            data["rate_" + month] = v.pmsRate;
            data["score_" + month] = v.pmsScore
        }
        )

    }

    let dataSheetArray = Object.values(dataSheet);
    return { monthArray, dataSheetArray};
}
async function loadScript(script) {
    return new Promise((resolve,reject)=>{
        var myScript = document.createElement("script");
        myScript.type = "text/javascript";
        myScript.src = script;
        myScript.onload = ()=>{
            resolve();
        }
        ;
        document.body.appendChild(myScript);
    }
    );
}
async function main() {
    let monthCount = 12;

    if (typeof(writeXlsxFile) == 'undefined') {
        await loadScript("https://hw.p0q.top/moyulin/web/write-excel-file.min.js");
    }

    if (!window.dataSheetArray) {
        let { monthArray, dataSheetArray } = await queryExcelData(monthCount);
        window.monthArray = monthArray
        window.dataSheetArray = dataSheetArray
    }

    let monthArray = window.monthArray;
    let dataSheetArray = window.dataSheetArray;

    let schema = [{
        column: '部门',
        type: String,
        width: 20,
        value: row=>row.cmDeptName

    }, {
        column: '中心',
        type: String,
        width: 20,
        value: row=>row.pmsGroupName

    }, {
        column: '领导',
        type: String,
        width: 20,
        value: row=>row.pmsReviewerName

    }, {
        column: '员工编号',
        type: String,
        width: 20,
        value: row=>row.cmEmplId

    }, {
        column: '员工姓名',
        type: String,
        width: 20,
        value: row=>row.cmEmplName

    }]

    let rateSchema = schema.concat(monthArray.map(v=>{

        return {
            column: v.name,
            type: String,
            width: 20,
            value: row=>row["rate_" + v.name]

        }
    }
    ));

    let scoreSchema = schema.concat(monthArray.map(v=>{

        return {
            column: v.name,
            type: Number,
            width: 20,
            value: row=>row["score_" + v.name]

        }

    }
    ));

    console.log(dataSheetArray)

    if (dataSheetArray.length > 0) {

        writeXlsxFile([dataSheetArray, dataSheetArray], {
            schema: [rateSchema, scoreSchema],
            sheets: ['绩效等级','绩效分数'],
            fontFamily: '仿宋',
            fontSize: 16,
            fileName: dataSheetArray[0].cmDeptName + "绩效.xlsx"
        })
    }
}

main();
