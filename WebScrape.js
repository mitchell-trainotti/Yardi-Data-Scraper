const ExportedData = [];
const goodPointers = [];
const executeLookups = async (pointers) => {
    ExportedData.push(['Pointer', 'Property Code', 'Intercom Code', 'Front Entrance Code', 'LockBox Code', 'Boiler Room Access', 'Master Code', 'Lockbox Location 1', 'Lockbox Code 1', 'Lockbox Location 2', 'Lockbox Code 2', 'Lockbox Location 3', 'Lockbox Code 3', 'Notes']);
    for (pointer of pointers) {
        url = `/67931ballast/Pages/SysUserData.aspx?sTableName=Propbut39&hPointer=${pointer}&Id=&iDisplayType=1&sCustomBtnType=PROPERTY%20BUTTONS&sTitle=Lock%20Box&caption=Lock%20Box`;

        const parser = new DOMParser();
        console.log(`Fetching: ${url}`)
        const res = await fetch(url);
        console.log(`Parsing html!`)
        const text = await res.text();
        var el = parser.parseFromString(text, "text/html");

        var propertyCode = getVal("bal_lock_propscode_TextBox", el, pointer);
        var intercomCode = getVal("bal_lock_intercomcode_TextBox", el, pointer);
        var frontEntranceAccess = getVal("bal_lock_FrontEntranceAccess_TextBox", el, pointer);
        var lockBoxCode = getVal("bal_lock_LockBoxCode_TextBox", el, pointer);
        var boilerRoomCode = getVal("bal_lock_BoilerRoomCode_TextBox", el, pointer);
        var masterCode = getVal("bal_lock_mastercode_TextBox", el, pointer);
        var box1Location = getVal("bal_lock_box1location_TextBox", el, pointer);
        var box1Code = getVal("bal_lock_box1code_TextBox", el, pointer);
        var box2Location = getVal("bal_lock_box2location_TextBox", el, pointer);
        var box2Code = getVal("bal_lock_box2code_TextBox", el, pointer);
        var box3Location = getVal("bal_lock_box3location_TextBox", el, pointer);
        var box3Code = getVal("bal_lock_box3code_TextBox", el, pointer);
        var notes = getVal("bal_lock_notes_TextBox", el, pointer);
        var AccessInfo = [pointer, propertyCode, intercomCode, frontEntranceAccess, lockBoxCode, boilerRoomCode, masterCode, box1Location, box1Code, box2Location, box2Code, box3Location, box3Code, notes];

        ExportedData.push(AccessInfo);
    }
    console.log("Writing to Excel...")
    exportToCsv(ExportedData);
};

exportToCsv = function (ExportedData) {
    var fileName = ("Access Info.csv");
    var CsvString = "";
    ExportedData.forEach(function (RowItem, RowIndex) {
        RowItem.forEach(function (ColItem, ColIndex) {
            CsvString += ColItem + ',';
        });
        CsvString += "\r\n";
    });
    CsvString = "data:application/csv," + encodeURIComponent(CsvString);
    var x = document.createElement("A");
    x.setAttribute("href", CsvString);
    x.setAttribute("download", fileName);
    document.body.appendChild(x);
    x.click();
}

function getVal (id, el, pointer){
    if (el.getElementById(id) != null){
        var propValue = el.getElementById(id).value;
        propValue = propValue.replace(/(\r\n|\n|\r|,|=)/gm, " ");
        if (String(propValue).charAt(0) == "0"){
            propValue = "_" + String(propValue);
        }
    }
    else{
        console.log("Null Pointer: " + pointer);
        var propValue = "N/A"
    }
    return propValue;   
}

const Pointers = ["98", "99", "100", "101", "102", "103", "108", "109", "110", "111", "112", "115", "118", "120", "121", "122", "123", "124", "125", "126", "127", "129", "130", "131", "132", "134", "148", "151", "153", "161", "162", "163", "164", "165", "166", "167", "168", "169", "170", "171", "172", "173", "174", "175", "176", "177", "178", "179", "180", "181", "182", "183", "184", "185", "186", "238", "239", "252", "253", "254", "255", "256", "257", "258", "259", "260", "261", "262", "263", "264", "265", "266", "267", "268", "269", "270", "271", "272", "274", "275", "277", "278", "279", "280", "282", "283", "284", "285", "286", "287", "312", "315", "316", "317", "320", "321", "322", "330", "331", "332", "333", "334", "335", "336", "339", "342", "343", "344", "348", "349", "350", "351", "352", "353", "354", "355", "356", "357", "358", "548", "599", "608", "611", "613", "615"]
var ignore = await executeLookups(Pointers);