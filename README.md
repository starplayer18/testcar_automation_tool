# testcar_automation_tool ver.0.1

현대자동차에서 시험차 담당자들의 전표처리를 위한 automation tool입니다.

<file>
- testcar_automation_tool.py
- config.json
- view.ui

* config.json
{
    "basic_info" : {
        "team_name": "",      #집행부서
        "budget_num": "",     #예산 번호
        "budget_name": "",    #예산명
        "date": "",           #의뢰일 "" 상태이면 사용하는 날짜가 자동으로 입력됨
        "purpose": "",        #용도
        "project_name": "",   #프로잭트명
        "manager": ""         #담당자
    },    
    "car_info" : {            #집행부서에서 운용하고 있는 차량 정보 / 차량번호:차량Code 형태로 추가 필요
        "1885": "JSN",
        "1888": "LX2",
        "1889": "IG PE",
        "1893": "RG3",
        "2472": "JW",
        "2910": "JK",
        "4959": "JK",
        "7156": "RG3",
        "7547": "JKe",
        "7548": "JKe",
        "7549": "RG3",
        "7550": "RG3",
        "7868": "RG3",
        "9009": "JK"
    }
}

시험차 전표처리를 위한 사전 정보를 config.json에 추가해야 합니다.


* testcar_automation_tool.py
config.json 내용을 parsing하여 필요한 정보를 set하고, view.ui를 통하여 사용자가 입력한 정보를 가공하여
구매의뢰 excel file / 시스템 입력용 excel file 2개를 export한다.
