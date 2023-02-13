function discription() {
    var urlAvrIndex = "https://docs.google.com/spreadsheets/d/1RkmT145G-KQcsSw0YTKjsH5cdfxwRjsyscBdblOJiOc/edit#gid=115958889";
    var ssAvrIndex = SpreadsheetApp.openByUrl(urlAvrIndex);
    var wsAvrIndex = ssAvrIndex.getSheetByName("Data");

    var wsDiscription = ss.getSheetByName("Discription");

    var analogs = getAnalogs(ss.getSheetByName("Аналоги"));
    var timers = getTimers(ss.getSheetByName("Лист3"));

    var data = wsAvrIndex.getDataRange().getValues();

    var arr = {};

    var htmlView = true;

    for (var i = 1; i < data.length; i++) {
    //for (var i = 152; i < 173; i++) {
        var schema = data[i][0];
        var article = data[i][1];
        var name = data[i][2];
        var plc = data[i][6];
        var modul = data[i][7];

        var brand = data[i][3];
        var type = data[i][4];
        var timeV = data[i][5];

        if (arr[article] == null) {
            var rec = {};

            rec.schema = schema;
            rec.article = article;
            rec.name = name;
            rec.plc = plc;

            if (modul !== "" && modul !== null) {
                rec.modul = modul;
            }

            rec.brands = {};

            var t = {};
            t[type] = timeV;
            rec.brands[brand] = t;

            arr[article] = rec;
        }
        else {
            var rec = arr[article];
            var brands = rec.brands;

            var t = {};

            if (brands[brand] != null) {
                t = brands[brand];
            }

            t[type] = timeV;
            brands[brand] = t;

            rec.brands = brands;

            arr[article] = rec;
        }
    }

    var plc_name = {};
    plc_name["197215"] = "EASY-E4-AC-12RC1";
    plc_name["197221"] = "EASY-E4-AC-8RE1";
    plc_name["232186"] = "EASY202-RE";
    plc_name["274104"] = "EASY512-AC-RC";
    plc_name["274115"] = "EASY719-AC-RC";

    Object.keys(arr).forEach(function (key, index) {
        var article = key;
        var rec = arr[article];
        var name = rec.name;
        var schema = rec.schema;

        var name1 = name.split(' ');
        var name1 = name1[0] + " " + name1[1] + " " + name1[2] + " " + name1[3];

        var str = "";

        if(htmlView){
            str = "<p>";
        }

        str += name1 + " ";
        str += "запрограммирован для управления АВР (автоматическим вводом резерва) по схеме ";
        str += schema + " на базе ";

        if (name.endsWith("для схем на контакторах")) {
            str += "контакторов.";
        } else {
            str += "автоматических выключателей:";
        }

        if(htmlView){
            str += "<br>";
        }else{
            str += "\n";
        }

        var i = 1;

        if (name.endsWith("для схем на авт. выкл.")) {
            var brands = {};
            Object.keys(rec.brands).sort().forEach(function (key) {
                brands[key] = rec.brands[key];
            });

            Object.keys(brands).forEach(function (key, index) {
                var brand = key;

                var types = {};
                Object.keys(brands[brand]).sort().forEach(function (key) {
                    types[key] = brands[brand][key];
                });

                Object.keys(types).forEach(function (key, index) {
                    var type = key;
                    var timeV = types[type];

                    str += i++ + ". " + brand + " " + type + " | время взвода пружины: " + timeV;
                    
                    if(htmlView){
                        str += "<br>";
                    }else{
                        str += "\n";
                    }
                });
            });

            if (timers[article] != null){
                str += "При наладке необходимо изменить уставки таймеров времени взвода пружин (" + timers[article] + 
                    ') в меню контроллера. Инструкция находится в разделе "Установка времени срабатывания таймеров" общего описания АВР.';
            }

            if(htmlView){
                str += "<br>";
            }else{
                str += "\n";
            }
        };

        if(htmlView){
            str += "<br>";
        }else{
            str += "\n";
        } 

        str += "Решение " + name1 + " (" + article + ") выполнено на базе свободно программируемого логического реле производства Eaton: " + rec.plc + " ";
        str += plc_name[rec.plc];

        if (rec.modul != null) {
            str += " + модуль расширения " + rec.modul + " " + plc_name[rec.modul];
        }

        if(htmlView){
            str += ".<br>";
        }else{
            str += ".\n";
        }

        if(analogs[article] != null){
            str += "Является аналогом контроллера " + analogs[article].analog + " " + analogs[article].type;

            if(htmlView){
                str += ".<br>";
            }else{
                str += ".\n";
            }
        }

        if(htmlView){
            str += "<br>";
        }else{
            str += "\n";
        }        

        str += "Для правильного выбора контроллера АВР по типам схемы и силовых аппаратов, и скачивания документации можно воспользоваться ";
        str += "<a href='https://sites.google.com/view/controlleravr/' target='_blank'>онлайн сервисом</a>" + ".";

        if(htmlView){
            str += "</p>";
        }

        wsDiscription.appendRow([article, str]);
    });
}

function getAnalogs(ws){
    var data = ws.getDataRange().getValues();

    var res = {};

    for (var i = 0; i < data.length; i++){
        var obj = {};
        obj.analog = data[i][1];
        obj.type = data[i][2];
        res[data[i][0]] = obj;
    }

    return res;
}

function getTimers(ws){
    var data = ws.getDataRange().getValues();

    var res = {};

    for (var i = 0; i < data.length; i++){
        if (data[i][2] != null && data[i][2] != ""){
            res[data[i][0]] = data[i][2];
        }
    }

    return res;
}