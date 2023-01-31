/**
 * @NApiVersion 2.x
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 */
define(['N/redirect', 'N/search', 'N/log', 'N/file', 'N/url', 'N/record', 'N/format', 'N/task', 'N/runtime'],
    /**
     * @param{redirect} redirect
     * @param{search} search
     * @param{log} log
     * @param{file} file
     * @param{url} url
     * @param{record} record
     * @param{format} format
     * @param{task} task
     * @param{runtime} runtime
     */
    function (redirect, search, log, file, url, record, format, task, runtime) {

        var params = null;

        /**
         * Definition of the Suitelet script trigger point.
         *
         * @param {Object} context
         * @param {ServerRequest} context.request - Encapsulation of the incoming request
         * @param {ServerResponse} context.response - Encapsulation of the Suitelet response
         * @Since 2015.2
         */
        function onRequest(context) {

            try {
                var response = context.response;
                var request = context.request;
                params = request.parameters;

                log.audit('Params SL', params);

                if (params.actionsl === 'del') {
                    deleteRol(params.idrol);
                } else {

                    var dataProcess = {}, dataArray = [];

                    var getDataRol = getDataFromRol(params.idrol);
                    var listaSello = getListaSelloPostal();
                    log.audit({ title: 'Length in search rols', details: getDataRol.length });
                    log.audit({ title: 'listaSello', details: listaSello });
                    var detailsArray = [];
                    for (var i = 0; i < getDataRol.length; i++) {
                        detailsArray.push(getDataRol[i].detailContract);
                    }
                    log.audit({ title: 'Action', details: 'Get Data' });
                    log.audit({
                        title: 'Data for getInfo', details: {
                            period: getDataRol[0].period,
                            month: getDataRol[0].month,
                            prefix: getDataRol[0].prefix,
                            dateRol: getDataRol[0].date_rol
                        }
                    });
                    var detailData = getAllDeailContract_v2(
                        detailsArray,
                        getDataRol[0].period,
                        getDataRol[0].month,
                        getDataRol[0].prefix,
                        getDataRol[0].date_rol,
                        getDataRol[0].weight,
                        listaSello);

                    for (var i = 0; i < detailData.length; i++) {
                        if (detailData[i].custrecord_efx_ms_con_ship_cus === "A06561") {
                            log.audit({ title: 'detailData', details: detailData[i] });
                        }
                    }
                    log.audit({
                        title: 'gobernances use getAllDeailContract_v2',
                        details: runtime.getCurrentScript().getRemainingUsage()
                    });
                    for (var j = 0; j < detailData.length; j++) {
                        dataArray.push(detailData[j]);
                    }
                    log.audit({ title: 'Sort', details: 'sort data' });
                    /**
                     * Orden y agrupamiento de datos
                     */
                    var newArraySort = setSortData(dataArray);
                    log.audit({ title: 'Sort', details: 'Finish sort data' });
                    log.audit({ title: 'Data 0', details: newArraySort[0] });
                    switch (params.actionsl) {
                        /** Archivo de Rol */
                        case 'e1':
                            var dataCSV = [];
                            var cantidad_alt = 0;
                            var peso_alt = 0;
                            var periodo = null;
                            for (var i = 0; i < newArraySort.length; i++) {
                                if (newArraySort[i].custrecord_efx_ms_con_period && newArraySort[i].custrecord_efx_ms_con_period === "4" || newArraySort[i].custrecord_efx_ms_con_period === 4) {
                                    cantidad_alt += newArraySort[i].custrecord_efx_ms_con_qty * 1;
                                    peso_alt = newArraySort[i].custrecord_efx_ms_con_w / newArraySort[i].custrecord_efx_ms_con_qty;
                                    periodo = newArraySort[i].custrecord_efx_ms_con_period;
                                }
                            }
                            if (periodo && (periodo === "4" || periodo === 4)) {
                                var contractDetail = {};
                                for (var i = 0; i < newArraySort.length; i++) {
                                    if (!contractDetail.hasOwnProperty(newArraySort[i].custrecord_efx_ms_con)) {
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con] = {};
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id] = {
                                            producto: "",
                                            peso: "",
                                            peso_total: ""
                                        };
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].publicacion = newArraySort[i].dateRol;
                                        if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].producto === "") {
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].producto = newArraySort[i].custrecord_efx_ms_con_ssi;
                                        }
                                        if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso === "") {
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso = newArraySort[i].w_ssi;
                                        }
                                        if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso_total === "") {
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso_total = (newArraySort[i].custrecord_efx_ms_con_qty * newArraySort[i].w_ssi);
                                        }
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].cantidad = newArraySort[i].custrecord_efx_ms_con_qty;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].tercero = (newArraySort[i].custrecord_efx_ms_con_contact) ? 'Si' : 'No';
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].usuario = newArraySort[i].custrecord_efx_ms_con_ship_cus;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].nombre = newArraySort[i].formula_nombre;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].contacto = newArraySort[i].custrecord_efx_ms_con_ship_cont;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].direccion = newArraySort[i].custrecord_efx_ms_con_cust_addr;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].compania = newArraySort[i].custrecord_efx_ms_con_sm;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].activo = (newArraySort[i].custrecord_efx_ms_con_act) ? 'Si' : 'No';
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].fecha_creacion = newArraySort[i].dateRol;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].creado_por = newArraySort[i].entityid;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].fecha_actualizado = newArraySort[i].lastmodified;
                                    } else {
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id] = {
                                            producto: "",
                                            peso: "",
                                            peso_total: ""
                                        };
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].publicacion = newArraySort[i].dateRol;
                                        if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].producto === "") {
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].producto = newArraySort[i].custrecord_efx_ms_con_ssi;
                                        }
                                        if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso === "") {
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso = newArraySort[i].w_ssi;
                                        }
                                        if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso_total === "") {
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso_total = (newArraySort[i].custrecord_efx_ms_con_qty * newArraySort[i].w_ssi);
                                        }
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].cantidad = newArraySort[i].custrecord_efx_ms_con_qty;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].tercero = (newArraySort[i].custrecord_efx_ms_con_contact) ? 'Si' : 'No';
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].usuario = newArraySort[i].custrecord_efx_ms_con_ship_cus;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].nombre = newArraySort[i].formula_nombre;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].contacto = newArraySort[i].custrecord_efx_ms_con_ship_cont;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].direccion = newArraySort[i].custrecord_efx_ms_con_cust_addr;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].compania = newArraySort[i].custrecord_efx_ms_con_sm;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].activo = (newArraySort[i].custrecord_efx_ms_con_act) ? 'Si' : 'No';
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].fecha_creacion = newArraySort[i].dateRol;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].creado_por = newArraySort[i].entityid;
                                        contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].fecha_actualizado = newArraySort[i].lastmodified;
                                    }
                                }
                                var contractDetailKey = Object.keys(contractDetail);
                                for (var key in contractDetailKey) {
                                    var keyIDS = Object.keys(contractDetail[contractDetailKey[key]]);
                                    for (var id in keyIDS) {
                                        dataCSV.push({
                                            'Publicacion': contractDetail[contractDetailKey[key]][keyIDS[id]].publicacion,
                                            'Producto': contractDetail[contractDetailKey[key]][keyIDS[id]].producto,
                                            'Cantidad': contractDetail[contractDetailKey[key]][keyIDS[id]].cantidad,
                                            'Peso': contractDetail[contractDetailKey[key]][keyIDS[id]].peso,
                                            'Peso Total': contractDetail[contractDetailKey[key]][keyIDS[id]].peso_total,
                                            'Tercero': contractDetail[contractDetailKey[key]][keyIDS[id]].tercero,
                                            'Usuario': contractDetail[contractDetailKey[key]][keyIDS[id]].usuario,
                                            'Nombre': contractDetail[contractDetailKey[key]][keyIDS[id]].nombre,
                                            'Contacto': contractDetail[contractDetailKey[key]][keyIDS[id]].contacto,
                                            'Direccion': contractDetail[contractDetailKey[key]][keyIDS[id]].direccion,
                                            'Compania': contractDetail[contractDetailKey[key]][keyIDS[id]].compania,
                                            'Activo': contractDetail[contractDetailKey[key]][keyIDS[id]].activo,
                                            'Fecha de creacion': contractDetail[contractDetailKey[key]][keyIDS[id]].fecha_creacion,
                                            'Creado por': contractDetail[contractDetailKey[key]][keyIDS[id]].creado_por,
                                            'Fecha de actualizacion': contractDetail[contractDetailKey[key]][keyIDS[id]].fecha_actualizado
                                        })
                                    }
                                }
                            } else {
                                for (var i = 0; i < newArraySort.length; i++) {
                                    dataCSV.push({
                                        'Publicacion': newArraySort[i].dateRol,
                                        'Producto': newArraySort[i].custrecord_efx_ms_con_ssi,
                                        'Cantidad': newArraySort[i].custrecord_efx_ms_con_qty,
                                        'Peso': newArraySort[i].custrecord_efx_ms_con_ssi_w,
                                        'Peso Total': (newArraySort[i].custrecord_efx_ms_con_ssi_w * newArraySort[i].custrecord_efx_ms_con_qty),
                                        'Tercero': (newArraySort[i].custrecord_efx_ms_con_contact) ? 'Si' : 'No',
                                        'Usuario': newArraySort[i].custrecord_efx_ms_con_ship_cus,
                                        'Nombre': newArraySort[i].formula_nombre,
                                        'Contacto': newArraySort[i].custrecord_efx_ms_con_ship_cont,
                                        'Direccion': newArraySort[i].custrecord_efx_ms_con_cust_addr,
                                        'Compania': newArraySort[i].custrecord_efx_ms_con_sm,
                                        'Activo': (newArraySort[i].custrecord_efx_ms_con_act) ? 'Si' : 'No',
                                        'Fecha de creacion': newArraySort[i].dateRol,
                                        'Creado por': newArraySort[i].entityid,
                                        'Fecha de actualizacion': newArraySort[i].lastmodified
                                    });
                                }
                            }
                            log.audit({ title: 'Details final in csvFile', details: dataCSV.length });
                            // context.response.write({output: JSON.stringify(dataCSV)})
                            var csvFile = generateExcel1(JSON.stringify(dataCSV));
                            if (csvFile) {
                                var urlFile = getFileURL(csvFile);
                                if (urlFile) {
                                    redirect.redirect({
                                        url: urlFile
                                    });
                                }
                            }
                            break;
                        /** Archivo de Consolidado */
                        case 'e3':
                            var dataCSV = [];
                            var cantidad_alt = 0;
                            var peso_alt = 0;
                            var periodo = null;
                            for (var i = 0; i < newArraySort.length; i++) {
                                if (newArraySort[i].custrecord_efx_ms_con_period === 4 || newArraySort[i].custrecord_efx_ms_con_period === "4") {
                                    cantidad_alt += newArraySort[i].custrecord_efx_ms_con_qty;
                                    peso_alt = newArraySort[i].custrecord_efx_ms_con_w / newArraySort[i].custrecord_efx_ms_con_qty;
                                    periodo = newArraySort[i].custrecord_efx_ms_con_period;
                                }
                                if (periodo && (periodo === "4" || periodo === 4)) {
                                    var contractDetail = {};
                                    var altI = 0;
                                    for (var i = 0; i < newArraySort.length; i++) {
                                        if (!contractDetail.hasOwnProperty(newArraySort[i].custrecord_efx_ms_con)) {
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con] = {};
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id] = {
                                                producto: "",
                                                peso: "",
                                                peso_total: ""
                                            };
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].organizacion = 'CAT';
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].no_documento = newArraySort[i].prefix + '-' + altI;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].tercero = (newArraySort[i].custrecord_efx_ms_con_contact) ? 'Si' : 'No';
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].usuario = newArraySort[i].custrecord_efx_ms_con_ship_cus;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].nombre = newArraySort[i].formula_nombre;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].contacto = newArraySort[i].custrecord_efx_ms_con_ship_cont;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].direccion = newArraySort[i].custrecord_efx_ms_con_cust_addr;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].referencia = newArraySort[i].custrecord_efx_ms_con_des;
                                            if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso === "") {
                                                contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso = newArraySort[i].w_ssi;
                                            }
                                            if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso_total === "") {
                                                contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso_total = (newArraySort[i].custrecord_efx_ms_con_qty * newArraySort[i].w_ssi);
                                            }
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].guia = '';
                                            altI++;
                                        } else {
                                            if (!contractDetail[newArraySort[i].custrecord_efx_ms_con].hasOwnProperty(newArraySort[i].id)) {
                                                altI++;
                                            }
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id] = {
                                                producto: "",
                                                peso: "",
                                                peso_total: ""
                                            };
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].organizacion = 'CAT';
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].no_documento = newArraySort[i].prefix + '-' + altI;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].tercero = (newArraySort[i].custrecord_efx_ms_con_contact) ? 'Si' : 'No';
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].usuario = newArraySort[i].custrecord_efx_ms_con_ship_cus;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].nombre = newArraySort[i].formula_nombre;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].contacto = newArraySort[i].custrecord_efx_ms_con_ship_cont;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].direccion = newArraySort[i].custrecord_efx_ms_con_cust_addr;
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].referencia = newArraySort[i].custrecord_efx_ms_con_des;
                                            if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso === "") {
                                                contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso = newArraySort[i].w_ssi;
                                            }
                                            if (contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso_total === "") {
                                                contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].peso_total = (newArraySort[i].custrecord_efx_ms_con_qty * newArraySort[i].w_ssi);
                                            }
                                            contractDetail[newArraySort[i].custrecord_efx_ms_con][newArraySort[i].id].guia = '';
                                        }
                                    }
                                    var contractDetailKey = Object.keys(contractDetail);
                                    for (var key in contractDetailKey) {
                                        var keyIDS = Object.keys(contractDetail[contractDetailKey[key]]);
                                        for (var id in keyIDS) {
                                            dataCSV.push({
                                                'Organizacion': contractDetail[contractDetailKey[key]][keyIDS[id]].organizacion,
                                                'No. Documento': contractDetail[contractDetailKey[key]][keyIDS[id]].no_documento,
                                                'Tercero': contractDetail[contractDetailKey[key]][keyIDS[id]].tercero,
                                                'Usuario': contractDetail[contractDetailKey[key]][keyIDS[id]].usuario,
                                                'Nombre': contractDetail[contractDetailKey[key]][keyIDS[id]].nombre,
                                                'Contacto': contractDetail[contractDetailKey[key]][keyIDS[id]].contacto,
                                                'Direccion': contractDetail[contractDetailKey[key]][keyIDS[id]].direccion,
                                                'Referencia': contractDetail[contractDetailKey[key]][keyIDS[id]].referencia,
                                                'Peso': contractDetail[contractDetailKey[key]][keyIDS[id]].peso,
                                                'Peso Total': contractDetail[contractDetailKey[key]][keyIDS[id]].peso_total,
                                                'Guia': contractDetail[contractDetailKey[key]][keyIDS[id]].guia
                                            })
                                        }
                                    }
                                } else {
                                    var peso = 0, peso_total = 0;
                                    peso = newArraySort[i].custrecord_efx_ms_con_w; //(newArraySort[i].custrecord_efx_ms_con_w / newArraySort[i].custrecord_efx_ms_con_qty);
                                    peso_total = newArraySort[i].custrecord_efx_ms_con_w * newArraySort[i].custrecord_efx_ms_con_qty;
                                    dataCSV.push({
                                        'Organizacion': 'CAT',
                                        'No. documento': newArraySort[i].prefix + '-' + (i + 1),
                                        'Tercero': (newArraySort[i].custrecord_efx_ms_con_contact) ? 'Si' : 'No',
                                        'Usuario': newArraySort[i].custrecord_efx_ms_con_ship_cus,
                                        'Nombre': newArraySort[i].formula_nombre,
                                        'Contacto': newArraySort[i].custrecord_efx_ms_con_ship_cont,
                                        'Direccion': newArraySort[i].custrecord_efx_ms_con_cust_addr,
                                        'Referencia': newArraySort[i].custrecord_efx_ms_con_des,
                                        'Peso': peso,
                                        'Peso Total': peso_total,
                                        'Guia': ''
                                    });
                                }
                            }
                            log.audit({ title: 'Details final in csvFile', details: dataCSV.length });
                            // context.response.write({output: JSON.stringify(dataCSV)})
                            var csvFile = generateExcel3(JSON.stringify(dataCSV));
                            if (csvFile) {
                                var urlFile = getFileURL(csvFile);
                                if (urlFile) {
                                    redirect.redirect({
                                        url: urlFile
                                    });
                                }
                            }
                            break;
                    }
                }
            } catch (e) {
                log.error('Error on onRequest', e);
            }

        }

        function deleteRol(idRol) {
            try {
                var taskObj = task.create({ taskType: task.TaskType.MAP_REDUCE });
                taskObj.scriptId = 'customscript_efx_ms_del_rol_mr';
                taskObj.deploymentId = 'customdeploy_efx_ms_del_rol_mr';
                taskObj.params = {
                    'custscript_efx_ms_rol_delete': idRol
                }
                var taskID = taskObj.submit();
                log.audit({ title: 'task for delete', details: taskID });
                if (taskID) {
                    record.submitFields({
                        type: 'customrecord_efx_ms_rol_env',
                        id: idRol,
                        values: {
                            'custrecord_efx_ms_re_task_delete': taskID,
                            'custrecord_efx_ms_re_status_delete': 0
                        }
                    });
                    redirect.toRecord({
                        type: 'customrecord_efx_ms_rol_env',
                        id: idRol,
                        parameters: {}
                    });
                }
            } catch (e) {
                log.error('Error on deleteRol', e);
            }
        }

        function getRolDetail(idRol) {
            try {
                var idDetailRol = [];
                var searchRol = search.create({
                    type: 'customrecord_efx_ms_detail_rol_env',
                    filters: [
                        ['custrecord_efx_ms_dre_re', search.Operator.IS, idRol],
                        'and',
                        ['isinactive', search.Operator.IS, 'F']
                    ],
                    columns: [
                        { name: 'internalid' }
                    ]
                });
                searchRol.run().each(function (res) {
                    idDetailRol.push(res.getValue({
                        name: 'internalid'
                    }));
                    return true;
                });
                return idDetailRol;
            } catch (e) {
                log.error('Error on getRolDetail', e);
            }
        }

        function getExcel1Data(idRol) {
            try {
                var data = [];
                var customrecord_efx_ms_detail_rol_envSearchObj = search.create({
                    type: "customrecord_efx_ms_detail_rol_env",
                    filters: [
                        ["custrecord_efx_ms_dre_re", search.Operator.IS, idRol]
                    ],
                    columns: [
                        { name: "internalid" },
                        { name: "created" },
                        { name: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_ssi", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_qty", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_w", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_unit_w", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_contact", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_ship_cus", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_addr", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_sm", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_act", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "created", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: 'entityid', join: 'owner' },
                        { name: "lastmodified", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: 'custrecord_efx_ms_con_ship_cont', join: 'custrecord_efx_ms_dre_detail_contract' },
                        { name: 'custrecord_efx_ms_con_period', join: 'custrecord_efx_ms_dre_detail_contract' }
                    ]
                });

                var cantidad_alt = 0;
                var period = null;
                var unit_w = null;
                customrecord_efx_ms_detail_rol_envSearchObj.run().each(function (result) {
                    var publicacion = result.getValue({ name: 'created' });
                    var producto = result.getText({
                        name: "custrecord_efx_ms_con_ssi",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var cantidad = result.getValue({
                        name: "custrecord_efx_ms_con_qty",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var peso = result.getValue({
                        name: "custrecord_efx_ms_con_w",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });

                    var periodo = result.getValue({
                        name: "custrecord_efx_ms_con_period",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });

                    var peso_total = result.getValue({
                        name: "custrecord_efx_ms_con_w",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var peso_unit = result.getValue({
                        name: "custrecord_efx_ms_con_unit_w",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });

                    if (periodo === "4") {
                        cantidad_alt += cantidad;
                        period = periodo;
                        unit_w = peso_unit;
                    }

                    var tercero = result.getValue({
                        name: "custrecord_efx_ms_con_contact",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var shipping_id = result.getValue({
                        name: "custrecord_efx_ms_con_sm",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var usuario = result.getText({
                        name: "custrecord_efx_ms_con_ship_cus",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var direccion = result.getValue({
                        name: "custrecord_efx_ms_con_addr",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    if (tercero) {
                        var related_add = result.getValue({
                            name: "custrecord_efx_ms_con_ship_cont", join: "custrecord_efx_ms_dre_detail_contract"
                        });
                        log.audit('related_add', related_add);
                        direccion = getFullAddress(related_add);
                    }
                    var compania = result.getText({
                        name: "custrecord_efx_ms_con_sm",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var activo = result.getValue({
                        name: "custrecord_efx_ms_con_act",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var fecha_creacion = result.getValue({
                        name: "created",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var creado_por = result.getValue({ name: 'entityid', join: 'owner' });
                    var actualizado = result.getValue({
                        name: "lastmodified",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    data.push({
                        publicacion: publicacion,
                        producto: producto,
                        cantidad: cantidad,
                        peso: (peso / cantidad) + ' ' + peso_unit,
                        peso_total: peso_total + ' ' + peso_unit,
                        tercero: (tercero) ? 'Si' : 'No',
                        shipping_id: shipping_id,
                        usuario: usuario,
                        direccion: direccion,
                        compania: compania,
                        activo: (activo) ? 'Si' : 'No',
                        fecha_creacion: fecha_creacion,
                        creado_por: creado_por,
                        actualizado: actualizado
                    });
                    return true;
                });
                if (period && period === "4") {
                    var unit = unit_w;
                    var peso_alt = data[0].peso / data[0].cantidad;
                    var peso_total = cantidad_alt * peso_alt;
                    for (var i = 0; i < data.length; i++) {
                        data[i].peso_total = peso_total;
                    }
                }
                return data
            } catch (e) {
                log.error('Error on getExcel1Data', e);
            }
        }

        function generateExcel1(dataRol) {
            try {
                // log.audit('Data Rol', dataRol);
                var json = JSON.parse(dataRol);
                var fields = Object.keys(json[0]);
                var replacer = function (key, value) {
                    return value === null ? '' : value
                };
                var csv = json.map(function (row) {
                    return fields.map(function (fieldName) {
                        return JSON.stringify(row[fieldName], replacer)
                    }).join(',')
                });
                csv.unshift(fields.join(','));
                csv = csv.join('\r\n');
                var fileObj = file.create({
                    name: 'rol_mensual.csv',
                    fileType: file.Type.CSV,
                    encoding: file.Encoding.WINDOWS_1252,
                    contents: csv
                });
                fileObj.folder = -15;
                var fileId = fileObj.save();


                return fileId;
            } catch (e) {
                log.error('Error on generateExcel1', e);
            }
        }

        function getExcel3Data(idRol) {
            try {
                var data = [];
                var customrecord_efx_ms_detail_rol_envSearchObj = search.create({
                    type: "customrecord_efx_ms_detail_rol_env",
                    filters: [
                        ["custrecord_efx_ms_dre_re", search.Operator.IS, idRol],
                        "AND",
                        ["custrecord_efx_ms_dre_detail_contract.custrecord_efx_ms_sal_ord_related", search.Operator.NONEOF, "@NONE@"],
                        "AND",
                        ["custrecord_efx_ms_dre_detail_contract.custrecord_efx_ms_con_sales_order_mirror", search.Operator.NONEOF, "@NONE@"]
                    ],
                    columns: [
                        { name: "internalid" },
                        { name: "created" },
                        { name: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_ssi", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_qty", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_w", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_unit_w", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_contact", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_ship_cus", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_addr", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_sm", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con_act", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "custrecord_efx_ms_con", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "created", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: "created", join: "custrecord_efx_ms_dre_re" },
                        { name: "custrecord_efx_ms_sal_ord_related", join: "custrecord_efx_ms_dre_detail_contract" },
                        {
                            name: "custrecord_efx_ms_con_sales_order_mirror",
                            join: "custrecord_efx_ms_dre_detail_contract"
                        },
                        { name: "custrecord_efx_ms_con_des", join: "custrecord_efx_ms_dre_detail_contract" },
                        { name: 'custrecord_efx_ms_con_ship_cont', join: 'custrecord_efx_ms_dre_detail_contract' }
                    ]
                });
                var countIndex = 1;
                customrecord_efx_ms_detail_rol_envSearchObj.run().each(function (result) {
                    var ordenRelacionada = result.getValue({
                        name: "custrecord_efx_ms_sal_ord_related",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var ordenEspejo = result.getValue({
                        name: "custrecord_efx_ms_con_sales_order_mirror",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var noRandom = new Date(format.parse({
                        value: result.getValue({ name: "created", join: "custrecord_efx_ms_dre_re" }),
                        type: format.Type.DATE
                    }));
                    var organizacion = 'CAT';
                    var n_documento = noRandom.getDate() + '' + (noRandom.getMonth() + 1) + '' + noRandom.getHours() + '' + noRandom.getMinutes() + '-' + countIndex;
                    var tercero = result.getValue({
                        name: "custrecord_efx_ms_con_contact",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var usuario = result.getText({
                        name: "custrecord_efx_ms_con_ship_cus",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var direccion = result.getValue({
                        name: "custrecord_efx_ms_con_addr",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    if (tercero === 'true') {
                        direccion = getFullAddress(result.getValue({
                            name: "custrecord_efx_ms_con_ship_cont", join: "custrecord_efx_ms_dre_detail_contract"
                        }));
                    }
                    var referencia = result.getValue({
                        name: "custrecord_efx_ms_con_des",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var peso = result.getValue({
                        name: "custrecord_efx_ms_con_w",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    }) + ' ' + result.getText({
                        name: "custrecord_efx_ms_con_unit_w",
                        join: "custrecord_efx_ms_dre_detail_contract"
                    });
                    var guia = '';

                    data.push({
                        organizacion: organizacion,
                        n_documento: n_documento,
                        // orden: result.getValue({name: "custrecord_efx_ms_dre_detail_contract"}),
                        tercero: (tercero) ? 'Si' : 'No',
                        usuario: usuario,
                        direccion: direccion,
                        referencia: referencia,
                        peso: peso,
                        guia: guia
                    });
                    countIndex++;
                    return true;
                });
                return data
            } catch (e) {
                log.error('Error on getExcel1Data', e);
            }
        }

        function generateExcel3(dataRol) {
            try {
                var json = JSON.parse(dataRol);
                var fields = Object.keys(json[0]);
                var replacer = function (key, value) {
                    return value === null ? '' : value
                };
                var csv = json.map(function (row) {
                    return fields.map(function (fieldName) {
                        return JSON.stringify(row[fieldName], replacer)
                    }).join(',')
                });
                csv.unshift(fields.join(','));
                csv = csv.join('\r\n');
                var fileObj = file.create({
                    name: 'rol_consolidado.csv',
                    fileType: file.Type.CSV,
                    encoding: file.Encoding.WINDOWS_1252,
                    contents: csv
                });
                fileObj.folder = -15;
                var fileId = fileObj.save();


                return fileId;
            } catch (e) {
                log.error('Error on generateExcel1', e);
            }
        }

        function getFileURL(fileId) {
            try {
                var fileObj = file.load({
                    id: fileId
                });

                var fileURL = fileObj.url;
                var scheme = 'https://';
                var host = url.resolveDomain({
                    hostType: url.HostType.APPLICATION
                });

                var urlFinal = scheme + host + fileURL;
                return urlFinal;
            } catch (e) {
                log.error('Error on getFileURL', e);
            }
        }

        function getDataFromRol(idRol) {
            try {
                var resultData = [];
                var customrecord_efx_ms_detail_rol_envSearchObj = search.create({
                    type: "customrecord_efx_ms_detail_rol_env",
                    filters:
                        [
                            ["custrecord_efx_ms_dre_re", search.Operator.IS, idRol]
                        ],
                    columns:
                        [
                            { name: "internalid" },
                            { name: "created" },
                            { name: "custrecord_efx_ms_dre_re" },
                            { name: "custrecord_efx_ms_dre_item" },
                            { name: "custrecord_efx_ms_dre_sm" },
                            { name: "custrecord_efx_ms_dre_detail_contract" },
                            {
                                name: "custrecord_efx_ms_con_period",
                                join: "CUSTRECORD_EFX_MS_DRE_DETAIL_CONTRACT"
                            },
                            {
                                name: 'custrecord_tkio_weight',
                                join: 'custrecord_efx_ms_dre_re',
                            },
                            {
                                name: "custrecord_efx_ms_con_sales_order_mirror",
                                join: "CUSTRECORD_EFX_MS_DRE_DETAIL_CONTRACT"
                            },
                            {
                                name: "formulanumeric",
                                formula: "TO_NUMBER(TO_CHAR({created},'MM'))"
                            },
                            {
                                name: "custrecord_efx_ms_re_period",
                                join: "custrecord_efx_ms_dre_re"
                            },
                            { name: "formulatext", formula: "TO_CHAR({created},'DDMMHHMI')" },
                            { name: "custrecord_efx_ms_re_year_apply", join: "custrecord_efx_ms_dre_re" }
                        ]
                });
                var myPagedResults = customrecord_efx_ms_detail_rol_envSearchObj.runPaged({
                    pageSize: 1000
                });
                var thePageRanges = myPagedResults.pageRanges;
                for (var i in thePageRanges) {
                    var thepageData = myPagedResults.fetch({
                        index: thePageRanges[i].index
                    });
                    thepageData.data.forEach(function (result) {
                        var itemService = result.getValue({ name: "custrecord_efx_ms_dre_item" });
                        var period = result.getValue({
                            name: "custrecord_efx_ms_con_period",
                            join: "CUSTRECORD_EFX_MS_DRE_DETAIL_CONTRACT"
                        });
                        var month = result.getValue({
                            name: "custrecord_efx_ms_re_period",
                            join: "custrecord_efx_ms_dre_re"
                        });
                        var weight = result.getValue({
                            name: 'custrecord_tkio_weight',
                            join: 'custrecord_efx_ms_dre_re',
                        })
                        var detailContract = result.getValue({ name: "custrecord_efx_ms_dre_detail_contract" });
                        var prefix = result.getValue({ name: "formulatext", formula: "TO_CHAR({created},'DDMMHHMI')" });
                        var date = format.parse({
                            value: result.getValue({ name: "created" }),
                            type: format.Type.DATE
                        });
                        var yearApplyRol = result.getValue({ name: "custrecord_efx_ms_re_year_apply", join: "custrecord_efx_ms_dre_re" })
                        var dateForRol = '01/'+((period<10)?'0'+period:period)+'/'+yearApplyRol;
                        resultData.push({
                            item: itemService,
                            period: period,
                            month: month,
                            detailContract: detailContract,
                            prefix: prefix,
                            date_rol: new Date(dateForRol),
                            weight: weight
                        });
                    });
                }
                return resultData;
            } catch (e) {
                log.error('Error on getDataFromRol', e);
            }
        }

        function getListaSelloPostal() {
            try {
                var listaSello = [];
                var customrecord_ms_pesos_sellosSearchObj = search.create({
                    type: "customrecord_ms_pesos_sellos",
                    filters: [],
                    columns:
                        [
                            { name: "custrecord_ms_sent_method" },
                            { name: "custrecord_ms_peso_min" },
                            { name: "custrecord_ms_peso_mx" },
                            { name: "custrecord_ms_sello_postal" },
                            { name: "custrecord_efx_ms_items" }
                        ]
                });
                var searchResult = customrecord_ms_pesos_sellosSearchObj.runPaged({
                    pageSize: 1000
                });

                var thePageRanges = searchResult.pageRanges;
                for (var i in thePageRanges) {
                    var thepageData = searchResult.fetch({
                        index: thePageRanges[i].index
                    });
                    thepageData.data.forEach(function (result) {
                        var pesoMin = result.getValue({ name: "custrecord_ms_peso_min" });
                        var pesoMax = result.getValue({ name: "custrecord_ms_peso_mx" });
                        var methodSent = result.getText({ name: "custrecord_ms_sent_method" });
                        var selloPostal = result.getValue({ name: "custrecord_ms_sello_postal" });
                        var articulo = (result.getText({ name: "custrecord_efx_ms_items" })).split(',');
                        listaSello.push({
                            articulo: articulo,
                            methodSent: methodSent,
                            selloPostal: selloPostal,
                            pesoMin: pesoMin,
                            pesoMax: pesoMax
                        });
                    })
                }
                return listaSello;
            } catch (e) {
                log.error('Error on getListaSelloPostal', e);
            }
        }

        function getInventoryItem(itemService, period, consecutive) {
            try {
                var filters = [];
                var itemResult = {};
                period = period * 1;
                consecutive = consecutive.toString();
                if (period === 1 || period === 2 || period === 3) {
                    filters = [
                        ["custrecord_efx_ms_item_sus", search.Operator.IS, itemService],
                        "AND",
                        ["custrecord_efx_ms_period_week.custrecord_efx_ms_period_month", search.Operator.IS, "T"],
                        "AND",
                        ["custrecord_efx_ms_period_week.custrecord_efx_ms_consecutive", search.Operator.IS, consecutive]
                    ];
                } else {
                    filters = [
                        ["custrecord_efx_ms_item_sus", search.Operator.ANYOF, itemService],
                        "AND",
                        ["custrecord_efx_ms_period_week.custrecord_efx_ms_period_month", search.Operator.IS, "F"],
                        "AND",
                        ["custrecord_efx_ms_period_week.custrecord_efx_ms_consecutive", search.Operator.IS, consecutive]
                    ];
                }
                var customrecord_efx_ms_sus_relatedSearchObj = search.create({
                    type: "customrecord_efx_ms_sus_related",
                    filters: filters,
                    columns:
                        [
                            { name: "custrecord_efx_ms_inventory_item" },
                            { name: "custrecord_efx_ms_item_sus" },
                            { name: "custrecord_efx_ms_period_week" },
                            { name: "weight", join: "custrecord_efx_ms_inventory_item" }
                        ]
                });
                var searchResult = customrecord_efx_ms_sus_relatedSearchObj.run().getRange({
                    start: 0,
                    end: 1
                });
                for (var i = 0; i < searchResult.length; i++) {
                    var getItem = searchResult[i].getText({ name: "custrecord_efx_ms_inventory_item" });
                    var getPeriod = searchResult[i].getText({ name: "custrecord_efx_ms_period_week" });
                    var getWeight = searchResult[i].getValue({
                        name: "weight",
                        join: "custrecord_efx_ms_inventory_item"
                    }) || 0;
                    itemResult = { item: getItem, period: getPeriod, weight: getWeight };
                }
                return itemResult;
            } catch (e) {
                log.error('Error on getInventoryItem', e);
            }
        }

        function getItemSales(itemsObj, weight) {
            try {
                var itemInv = {};
                var month, itemSus, consecutives = [];
                var itemsKey = Object.keys(itemsObj);
                for (var i in itemsKey) {
                    itemSus = itemsKey[i];
                    month = (itemsObj[itemsKey[i]][0].monthly) ? "T" : "F";
                    for (var j in itemsObj[itemsKey[i]]) {
                        if (consecutives.indexOf(itemsObj[itemsKey[i]][j].consecutive) < 0) {
                            consecutives.push(parseInt(itemsObj[itemsKey[i]][j].consecutive))
                        }
                    }
                    var minimo = Math.min.apply(null, consecutives);
                    var limite = Math.max.apply(null, consecutives);
                    var filts = [
                        ["custrecord_efx_ms_period_week.custrecord_efx_ms_period_month", search.Operator.IS, month],
                        "AND",
                        ["custrecord_efx_ms_item_sus", search.Operator.ANYOF, itemSus],
                        /*"AND",
                        ["formulanumeric: TO_NUMBER({custrecord_efx_ms_period_week.custrecord_efx_ms_consecutive})", search.Operator.LESSTHANOREQUALTO, limite],
                        "AND",
                        ["formulanumeric: TO_NUMBER({custrecord_efx_ms_period_week.custrecord_efx_ms_consecutive})", search.Operator.GREATERTHANOREQUALTO, minimo]*/
                    ];
                    var searchObj = search.create({
                        type: "customrecord_efx_ms_sus_related",
                        filters: filts,
                        columns:
                            [
                                { name: "internalid" },
                                { name: "custrecord_efx_ms_item_sus" },
                                { name: "custrecord_efx_ms_inventory_item" },
                                {
                                    name: "custrecord_efx_ms_consecutive",
                                    join: "custrecord_efx_ms_period_week",
                                    sort: search.Sort.ASC
                                },
                                { name: "custrecord_efx_ms_year_apply", sort: search.Sort.ASC },
                                { name: "custrecord_efx_ms_period_week" },
                                { name: "weight", join: "custrecord_efx_ms_inventory_item" },
                                { name: "displayname", join: "custrecord_efx_ms_inventory_item" }
                            ]
                    });
                    var countData = searchObj.runPaged().count;
                    // log.audit({title: 'Count results items', details: countData});
                    searchObj.run().each(function (result) {
                        if (itemInv.hasOwnProperty(itemsKey[i])) {
                            var con = result.getValue({
                                name: "custrecord_efx_ms_consecutive",
                                join: "custrecord_efx_ms_period_week",
                                sort: search.Sort.ASC
                            });
                            var itm = result.getValue({ name: "custrecord_efx_ms_inventory_item" });
                            var yer = result.getValue({ name: "custrecord_efx_ms_year_apply", sort: search.Sort.ASC });
                            var getItem = result.getText({ name: "custrecord_efx_ms_inventory_item" });
                            var getPeriod = result.getText({ name: "custrecord_efx_ms_period_week" });
                            // var getWeight = weight;
                            var getWeight = (weight.length !== 0) ? weight : result.getValue({ name: "weight", join: "custrecord_efx_ms_inventory_item" }) || 0;

                            var nameInventory = result.getValue({ name: "displayname", join: "custrecord_efx_ms_inventory_item" });
                            itemInv[itemsKey[i]].push({
                                consecutive: con,
                                item: itm,
                                year: yer,
                                itemName: getItem,
                                period: getPeriod,
                                weight: getWeight,
                                nameInventory: nameInventory
                            });
                        } else {
                            itemInv[itemsKey[i]] = [];
                            var con = result.getValue({
                                name: "custrecord_efx_ms_consecutive",
                                join: "custrecord_efx_ms_period_week",
                                sort: search.Sort.ASC
                            });
                            var itm = result.getValue({ name: "custrecord_efx_ms_inventory_item" });
                            var yer = result.getValue({ name: "custrecord_efx_ms_year_apply", sort: search.Sort.ASC });
                            var getItem = result.getText({ name: "custrecord_efx_ms_inventory_item" });
                            var getPeriod = result.getText({ name: "custrecord_efx_ms_period_week" });
                            var getWeight = (weight.length !== 0) ? weight : result.getValue({ name: "weight", join: "custrecord_efx_ms_inventory_item" });
                            var nameInventory = result.getValue({ name: "displayname", join: "custrecord_efx_ms_inventory_item" });
                            itemInv[itemsKey[i]].push({
                                consecutive: con,
                                item: itm,
                                year: yer,
                                itemName: getItem,
                                period: getPeriod,
                                weight: getWeight,
                                nameInventory: nameInventory
                            });
                        }
                        return true;
                    });
                }
                return itemInv;
            } catch (e) {
                log.error('Error on getItemSales', e);
            }
        }

        function weekCount(year, month_number) {

            // month_number is in the range 1..12

            var firstOfMonth = new Date(year, month_number - 1, 1);
            var lastOfMonth = new Date(year, month_number, 0);

            var used = firstOfMonth.getDay() + lastOfMonth.getDate();

            return Math.ceil(used / 7);
        }

        function getWeek(d) {

            var oneJan =
                new Date(d.getFullYear(), 0, 1);

            // calculating number of days
            //in given year before given date

            var numberOfDays =
                Math.floor((d - oneJan) / (24 * 60 * 60 * 1000));

            // adding 1 since this.getDay()
            //returns value starting from 0

            return Math.ceil((d.getDay() + 1 + numberOfDays) / 7);
        }

        function getWeekNumber(d) {
            // Copy date so don't modify original
            d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
            // Set to nearest Thursday: current date + 4 - current day number
            // Make Sunday's day number 7
            d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
            // Get first day of year
            var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
            // Calculate full weeks to nearest Thursday
            var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
            // Return array of year and week number
            return weekNo;
        }

        function getFullAddress(idAddressRelated) {
            try {
                var addressSearchObj = search.create({
                    type: "customrecord_efx_ms_custom_address",
                    filters:
                        [
                            ["internalid", search.Operator.IS, idAddressRelated]
                        ],
                    columns:
                        [
                            { name: "custrecord_efx_ms_ca_customer" },
                            { name: "custrecord_efx_ms_ca_addr" }
                        ]
                });
                var searchResult = addressSearchObj.run().getRange({ start: 0, end: 1 });
                var addressText = "";
                for (var i = 0; i < searchResult.length; i++) {
                    addressText = searchResult[i].getValue({ name: 'custrecord_efx_ms_ca_addr' });
                }
                return addressText;
            } catch (e) {

            }
        }

        function getCustomerExtNumber(customersID) {
            try {
                var customerData = {};
                var customerKey = Object.keys(customersID), ids = [];
                var customersRepeat = [];
                for (var i in customerKey) {
                    ids.push(customersID[customerKey[i]].id);
                }
                var filt = [
                    ["internalid", search.Operator.ANYOF, ids]
                ]
                var customerSearchObj = search.create({
                    type: search.Type.CUSTOMER,
                    filters: filt,
                    columns:
                        [
                            { name: "custrecord_streetnum", join: "billingAddress" },
                            { name: "entityid" },
                            { name: "internalid" }
                        ]
                });
                var searchResultCount = customerSearchObj.runPaged().count;
                if (searchResultCount > 0) {
                    customerSearchObj.run().each(function (result) {
                        var streetnum = result.getValue({ name: "custrecord_streetnum", join: "billingAddress" });
                        var customer = result.getValue({ name: "entityid" });
                        var cust_id = result.getValue({ name: "internalid" });
                        if (!customerData.hasOwnProperty(cust_id)) {
                            customerData[cust_id] = {};
                            customerData[cust_id] = {
                                id: cust_id,
                                customer: customer,
                                streetnum: streetnum
                            };
                        } else {
                            customerData[cust_id] = {
                                id: cust_id,
                                customer: customer,
                                streetnum: streetnum
                            };
                        }
                        return true;
                    });
                }
                return customerData;
            } catch (e) {
                log.error({ title: 'Error on getCustomerExtNumber', details: e });
            }
        }

        function getAllDeailContract_v2(detailID, period, consecutive, prefix, dateRol, weight, listaSello) {
            try {
                var getFirstSunday = function (month, year) {
                    try {
                        var tempDate = new Date();
                        tempDate.setHours(0, 0, 0, 0);
                        tempDate.setMonth(month);
                        tempDate.setYear(year);
                        tempDate.setDate(1);

                        var day = tempDate.getDay();
                        var toNextSun = day !== 0 ? 7 - day : 0;
                        tempDate.setDate(tempDate.getDate() + toNextSun);

                        return tempDate.getDate();
                    } catch (e) {
                        log.error({ title: 'Error on getFirstSunday', details: e });
                    }
                }
                period = period * 1;
                var detailData = [];
                var searchDetail = search.create({
                    type: 'customrecord_efx_ms_contract_detail',
                    filters: [
                        ['internalid', search.Operator.ANYOF, detailID],
                        'AND',
                        ['isinactive', search.Operator.IS, 'F']
                    ],
                    columns: [
                        { name: 'internalid' },
                        { name: 'custrecord_efx_ms_con_postmark' },
                        {
                            name: 'formulatext_postmark',
                            formula: 'TO_CHAR({custrecord_efx_ms_con_postmark})',
                        },
                        { name: 'custrecord_efx_ms_con_sd' },
                        { name: 'custrecord_efx_ms_con_ed' },
                        { name: 'custrecord_efx_ms_con_contact' },
                        { name: 'custrecord_efx_ms_ca_addr', join: 'custrecord_efx_ms_con_ship_cont' },
                        {
                            name: "formulatext",
                            formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billzipcode}  END",
                            sort: search.Sort.ASC
                        },
                        { name: 'custrecord_efx_ms_con_ship_cus' },
                        { name: 'custrecord_efx_ms_con_addr' },
                        { name: 'custrecord_efx_ms_con_des' },
                        { name: 'custrecord_efx_ms_con_qty' },
                        { name: 'custrecord_efx_ms_con_w' },
                        { name: 'custrecord_efx_ms_con_unit_w' },
                        { name: 'custrecord_efx_ms_con_sm' },
                        { name: 'custrecord_efx_ms_con_ssi' },
                        { name: 'displayname', join: 'custrecord_efx_ms_con_ssi' },
                        { name: 'custitem_efx_ms_service_weight', join: 'custrecord_efx_ms_con_ssi' },
                        { name: 'custrecord_efx_ms_con' },
                        { name: 'custrecord_efx_ms_con_tp' },
                        {
                            name: "formulatext_state",
                            formula: "CASE WHEN{custrecord_efx_ms_con_contact} = 'T' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END WHEN{custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE{custrecord_efx_ms_con_ship_cus.billstate} END END"
                        },
                        {
                            name: "formulatext_cust_address",
                            formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billaddress1} || '' || '-' || '' || ' _ ' || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billaddress3} || '' || '-' || '' || {custrecord_efx_ms_con_des} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billaddress2} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billzipcode} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billcity} || '' || '-' || '' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE {custrecord_efx_ms_con_ship_cus.billstate} END || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billcountry} WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_street} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_number} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_int_number} || '' || '-' || '' || {custrecord_efx_ms_con_des} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_col} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_city} || '' || '-' || '' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_country} END"
                        },
                        { name: 'custrecord_efx_ms_con_act' },
                        { name: 'lastmodified' },
                        { name: 'entityid', join: 'owner' },
                        { name: 'custrecord_efx_ms_con' },
                        {
                            name: "formulatext_nombre",
                            formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'F' THEN {custrecord_efx_ms_con_ship_cus} || ' ' || {custrecord_efx_ms_con_ship_cus.companyname} WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'T' THEN {custrecord_efx_ms_con_ship_cus} || ' ' || {custrecord_efx_ms_con_ship_cus.firstname} || ' ' || {custrecord_efx_ms_con_ship_cus.lastname} END WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_atention} END"
                        }
                    ]
                });
                log.audit({ title: 'Filters in get data', details: searchDetail.filters });
                var itemDataObj = {};
                var customersObj = {};
                var idrepeat = [];
                log.audit({ title: 'result count contract detail', details: searchDetail.runPaged().count });
                var myPagedResults = searchDetail.runPaged({
                    pageSize: 1000
                });
                var thePageRanges = myPagedResults.pageRanges;
                for (var i in thePageRanges) {
                    var thepageData = myPagedResults.fetch({
                        index: thePageRanges[i].index
                    });
                    thepageData.data.forEach(function (result, index_val) {
                        if (period === 4 || period === "4") {
                            var d = new Date(dateRol);
                            d.setMonth(consecutive - 1);
                            d.setDate(getFirstSunday(consecutive - 1, d.getFullYear()));
                            var weekC = weekCount(d.getFullYear(), parseInt(consecutive) /*d.getMonth() + 1*/);
                            if (index_val === 0) {
                                log.audit({
                                    title: 'result in WeekCount', details: {
                                        year: d.getFullYear(),
                                        month: parseInt(consecutive),
                                        weeks_in_month: weekC
                                    }
                                });
                            }
                            var noWeek = null;
                            var fYear = null;
                            for (var j = 0; j < weekC; j++) {
                                var weeks_add = 7 * (j);
                                noWeek = new Date(d.getTime() + (24 * 60 * 60 * 1000 * weeks_add));
                                var conse = getWeek(new Date(noWeek.getFullYear(), noWeek.getMonth(), noWeek.getDate()));
                                if (index_val === 0) { log.audit({ title: 'no week', details: { 0: noWeek, 1: conse } }); }
                                fYear = noWeek.getFullYear();

                                var internal_id = result.getValue({ name: 'internalid' });
                                // var custrecord_efx_ms_con_postmark = result.getText({
                                //     name: 'custrecord_efx_ms_con_postmark'
                                // });
                                var custrecord_efx_ms_con_postmark = result.getValue({
                                    name: 'formulatext_postmark',
                                    formula:
                                        'TO_CHAR({custrecord_efx_ms_con_postmark})',
                                }) || '';
                                var custrecord_efx_ms_con_sd = result.getValue({
                                    name: 'custrecord_efx_ms_con_sd'
                                });
                                var custrecord_efx_ms_con_ed = result.getValue({
                                    name: 'custrecord_efx_ms_con_ed'
                                });
                                var contact = result.getValue({ name: 'custrecord_efx_ms_con_contact' });
                                var custrecord_efx_ms_con_contact = contact;
                                var custrecord_efx_ms_con_ship_cont = result.getValue({
                                    name: 'custrecord_efx_ms_ca_addr',
                                    join: 'custrecord_efx_ms_con_ship_cont'
                                });
                                var custrecord_efx_ms_con_ship_cus = result.getText({
                                    name: 'custrecord_efx_ms_con_ship_cus'
                                });
                                var con_ship_cus = result.getValue({
                                    name: 'custrecord_efx_ms_con_ship_cus'
                                });
                                if (custrecord_efx_ms_con_contact === false) {
                                    if (!customersObj.hasOwnProperty(con_ship_cus)) {
                                        customersObj[con_ship_cus] = {};
                                        customersObj[con_ship_cus] = {
                                            id: con_ship_cus,
                                            customer: custrecord_efx_ms_con_ship_cus
                                        };
                                    }
                                } else {
                                    idrepeat.push({
                                        id: con_ship_cus,
                                        customer: custrecord_efx_ms_con_ship_cus
                                    })
                                }
                                var custrecord_efx_ms_con_addr = result.getValue({
                                    name: 'custrecord_efx_ms_con_addr'
                                });
                                var custrecord_efx_ms_con_cust_addr = result.getValue({
                                    name: "formulatext_cust_address",
                                    formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billaddress1} || '' || '-' || '' || ' _ ' || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billaddress3} || '' || '-' || '' || {custrecord_efx_ms_con_des} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billaddress2} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billzipcode} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billcity} || '' || '-' || '' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE {custrecord_efx_ms_con_ship_cus.billstate} END || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billcountry} WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_street} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_number} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_int_number} || '' || '-' || '' || {custrecord_efx_ms_con_des} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_col} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_city} || '' || '-' || '' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_country} END"
                                });
                                var custrecord_efx_ms_con = result.getValue({ name: 'custrecord_efx_ms_con' });
                                var custrecord_efx_ms_con_des = result.getValue({ name: 'custrecord_efx_ms_con_des' });
                                var custrecord_efx_ms_con_qty = result.getValue({ name: 'custrecord_efx_ms_con_qty' });
                                var custrecord_efx_ms_con_w = (weight.length !== 0) ? weight : result.getValue({ name: 'custrecord_efx_ms_con_w' });
                                var custrecord_efx_ms_con_unit_w = result.getText({ name: 'custrecord_efx_ms_con_unit_w' });
                                var custrecord_efx_ms_con_sm = result.getText({ name: 'custrecord_efx_ms_con_sm' });
                                var paqueteria = custrecord_efx_ms_con_sm;
                                var methodSent = custrecord_efx_ms_con_sm;
                                var iOF = custrecord_efx_ms_con_sm.toUpperCase().indexOf('SEPOMEX ');
                                if (iOF >= 0) {
                                    paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                                    methodSent = custrecord_efx_ms_con_sm.replace(paqueteria, '');
                                    methodSent = methodSent.replace(' ', '');
                                }
                                var custrecord_efx_ms_con_tp = result.getText({ name: 'custrecord_efx_ms_con_tp' });
                                var itemSSI = result.getValue({ name: 'custrecord_efx_ms_con_ssi' }).toString();
                                var itemNameSSI = result.getValue({ name: 'displayname', join: 'custrecord_efx_ms_con_ssi' });
                                var monthly = true;
                                if (period === 4 || period === '4') {
                                    monthly = false;
                                }
                                if (itemDataObj.hasOwnProperty(itemSSI) === false) {
                                    itemDataObj[itemSSI] = [];
                                    itemDataObj[itemSSI].push({
                                        consecutive: conse,
                                        monthly: monthly,
                                        yearApply: fYear
                                    });
                                } else {
                                    itemDataObj[itemSSI].push({
                                        consecutive: conse,
                                        monthly: monthly,
                                        yearApply: fYear
                                    });
                                }
                                var itemData = '';
                                var custrecord_efx_ms_con_ssi = '';
                                var custrecord_efx_ms_con_ssi_w = '';
                                var formulatext_month = '';
                                var custrecord_efx_ms_con_period = period;
                                var cpCustom = result.getValue({
                                    name: "formulatext",
                                    formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billzipcode}  END",
                                    sort: search.Sort.ASC
                                });
                                var state = result.getValue({
                                    name: "formulatext_state",
                                    formula: "CASE WHEN{custrecord_efx_ms_con_contact} = 'T' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END WHEN{custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE{custrecord_efx_ms_con_ship_cus.billstate} END END"
                                });
                                var matchState = matchStateCode(state);
                                if (matchState != null) {
                                    state = matchState;
                                }
                                var w_ssi = (weight.length !== 0) ? weight : result.getValue({ name: 'custitem_efx_ms_service_weight', join: 'custrecord_efx_ms_con_ssi' });
                                // var w_ssi = weight;
                                var peso = (parseFloat(custrecord_efx_ms_con_qty) * parseFloat(custrecord_efx_ms_con_w));
                                if (weight.length !== 0) {
                                    for (var i = 0; i < listaSello.length; i++) {
                                        if (peso >= parseFloat(listaSello[i].pesoMin) && peso <= parseFloat(listaSello[i].pesoMax) && (listaSello[i].articulo).indexOf(itemNameSSI) !== -1 && (custrecord_efx_ms_con_postmark === listaSello[i].selloPostal)) {
                                            custrecord_efx_ms_con_sm = listaSello[i].methodSent;
                                            paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                                            custrecord_efx_ms_con_postmark = listaSello[i].selloPostal;
                                            break;
                                        }
                                    }
                                }
                                var custrecord_efx_ms_con_act = result.getValue({ name: 'custrecord_efx_ms_con_act' });
                                var lastmodified = result.getValue({ name: 'lastmodified' });
                                var entityid = result.getValue({ name: 'entityid', join: 'owner' });
                                var custrecord_efx_ms_con = result.getValue({ name: 'custrecord_efx_ms_con' });

                                var formula_nombre = result.getValue({
                                    name: "formulatext_nombre",
                                    formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'F' THEN {custrecord_efx_ms_con_ship_cus} || ' ' || {custrecord_efx_ms_con_ship_cus.companyname} WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'T' THEN {custrecord_efx_ms_con_ship_cus} || ' ' || {custrecord_efx_ms_con_ship_cus.firstname} || ' ' || {custrecord_efx_ms_con_ship_cus.lastname} END WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_atention} END"
                                });
                                detailData.push({
                                    id: internal_id,
                                    custrecord_efx_ms_con_postmark: custrecord_efx_ms_con_postmark,
                                    custrecord_efx_ms_con_sd: custrecord_efx_ms_con_sd,
                                    custrecord_efx_ms_con_ed: custrecord_efx_ms_con_ed,
                                    custrecord_efx_ms_con_contact: custrecord_efx_ms_con_contact,
                                    custrecord_efx_ms_con_ship_cont: custrecord_efx_ms_con_ship_cont,
                                    custrecord_efx_ms_con_ship_cus: custrecord_efx_ms_con_ship_cus,
                                    custrecord_efx_ms_con_addr: custrecord_efx_ms_con_addr,
                                    custrecord_efx_ms_con_cust_addr: custrecord_efx_ms_con_cust_addr,
                                    custrecord_efx_ms_con_des: custrecord_efx_ms_con_des,
                                    custrecord_efx_ms_con_qty: parseInt(custrecord_efx_ms_con_qty),
                                    custrecord_efx_ms_con_w: custrecord_efx_ms_con_w,
                                    custrecord_efx_ms_con_unit_w: custrecord_efx_ms_con_unit_w,
                                    custrecord_efx_ms_con_sm: custrecord_efx_ms_con_sm,
                                    custrecord_efx_ms_con_ssi: custrecord_efx_ms_con_ssi,
                                    custrecord_efx_ms_con_ssi_w: custrecord_efx_ms_con_ssi_w,
                                    formulatext_month: formulatext_month,
                                    prefix: prefix,
                                    cpCustom: cpCustom,
                                    state: state,
                                    custrecord_efx_ms_con: custrecord_efx_ms_con,
                                    custrecord_efx_ms_con_period: period,
                                    custrecord_efx_ms_con_tp: custrecord_efx_ms_con_tp,
                                    consecutive: conse.toString(),
                                    yearApply: fYear.toString(),
                                    itemSSI: itemSSI,
                                    iOF: iOF,
                                    paqueteria: paqueteria,
                                    methodSent: methodSent,
                                    peso: peso,
                                    con_ship_cus: con_ship_cus,
                                    itemNameSSI: itemNameSSI,
                                    w_ssi: w_ssi,
                                    custrecord_efx_ms_con_act: custrecord_efx_ms_con_act,
                                    lastmodified: lastmodified,
                                    entityid: entityid,
                                    dateRol: dateRol,
                                    custrecord_efx_ms_con: custrecord_efx_ms_con,
                                    formula_nombre: formula_nombre
                                });
                            }
                        } else {
                            var d2 = new Date(dateRol);
                            var internal_id = result.getValue({ name: 'internalid' });
                            var custrecord_efx_ms_con_postmark = result.getValue({
                                name: 'formulatext_postmark',
                                formula:
                                    'TO_CHAR({custrecord_efx_ms_con_postmark})',
                            }) || '';
                            var custrecord_efx_ms_con_sd = result.getValue({
                                name: 'custrecord_efx_ms_con_sd'
                            });
                            var custrecord_efx_ms_con_ed = result.getValue({
                                name: 'custrecord_efx_ms_con_ed'
                            });
                            var contact = result.getValue({ name: 'custrecord_efx_ms_con_contact' });
                            var custrecord_efx_ms_con_contact = contact;
                            var custrecord_efx_ms_con_ship_cont = result.getValue({
                                name: 'custrecord_efx_ms_ca_addr',
                                join: 'custrecord_efx_ms_con_ship_cont'
                            });
                            var custrecord_efx_ms_con_ship_cus = result.getText({
                                name: 'custrecord_efx_ms_con_ship_cus'
                            });
                            var con_ship_cus = result.getValue({
                                name: 'custrecord_efx_ms_con_ship_cus'
                            });
                            if (custrecord_efx_ms_con_contact === false) {
                                if (!customersObj.hasOwnProperty(con_ship_cus)) {
                                    customersObj[con_ship_cus] = {};
                                    customersObj[con_ship_cus] = {
                                        id: con_ship_cus,
                                        customer: custrecord_efx_ms_con_ship_cus
                                    };
                                }
                            } else {
                                idrepeat.push({
                                    id: con_ship_cus,
                                    customer: custrecord_efx_ms_con_ship_cus
                                })
                            }
                            var custrecord_efx_ms_con_addr = result.getValue({
                                name: 'custrecord_efx_ms_con_addr'
                            });
                            var custrecord_efx_ms_con_cust_addr = result.getValue({
                                name: "formulatext_cust_address",
                                formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billaddress1} || '' || '-' || '' || ' _ ' || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billaddress3} || '' || '-' || '' || {custrecord_efx_ms_con_des} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billaddress2} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billzipcode} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billcity} || '' || '-' || '' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE {custrecord_efx_ms_con_ship_cus.billstate} END || '' || '-' || '' || {custrecord_efx_ms_con_ship_cus.billcountry} WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_street} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_number} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_int_number} || '' || '-' || '' || {custrecord_efx_ms_con_des} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_col} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_city} || '' || '-' || '' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END || '' || '-' || '' || {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_country} END"
                            });
                            var custrecord_efx_ms_con = result.getValue({ name: 'custrecord_efx_ms_con' });
                            var custrecord_efx_ms_con_des = result.getValue({ name: 'custrecord_efx_ms_con_des' });
                            var custrecord_efx_ms_con_qty = result.getValue({ name: 'custrecord_efx_ms_con_qty' });
                            var custrecord_efx_ms_con_w = (weight.length !== 0) ? weight : result.getValue({ name: 'custrecord_efx_ms_con_w' });
                            var custrecord_efx_ms_con_unit_w = result.getText({ name: 'custrecord_efx_ms_con_unit_w' });
                            var custrecord_efx_ms_con_sm = result.getText({ name: 'custrecord_efx_ms_con_sm' });
                            var methodSent = custrecord_efx_ms_con_sm;
                            var paqueteria = custrecord_efx_ms_con_sm;
                            var iOF = custrecord_efx_ms_con_sm.toUpperCase().indexOf('SEPOMEX ');
                            if (iOF >= 0) {
                                paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                                methodSent = custrecord_efx_ms_con_sm.replace(paqueteria, '');
                                methodSent = methodSent.replace(' ', '');
                            }
                            var custrecord_efx_ms_con_tp = result.getText({ name: 'custrecord_efx_ms_con_tp' });
                            var itemSSI = result.getValue({ name: 'custrecord_efx_ms_con_ssi' }).toString();
                            var itemNameSSI = result.getValue({ name: 'displayname', join: 'custrecord_efx_ms_con_ssi' });
                            var monthly = true;
                            fYear = d2.getFullYear();
                            if (itemDataObj.hasOwnProperty(itemSSI) === false) {
                                itemDataObj[itemSSI] = [];
                                itemDataObj[itemSSI].push({
                                    consecutive: consecutive,
                                    monthly: monthly,
                                    yearApply: fYear
                                });
                            } else {
                                itemDataObj[itemSSI].push({
                                    consecutive: consecutive,
                                    monthly: monthly,
                                    yearApply: fYear
                                });
                            }
                            var custrecord_efx_ms_con_ssi = '';
                            var custrecord_efx_ms_con_ssi_w = '';
                            var formulatext_month = '';
                            var cpCustom = result.getValue({
                                name: "formulatext",
                                formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billzipcode}  END",
                                sort: search.Sort.ASC
                            });
                            var state = result.getValue({
                                name: "formulatext_state",
                                formula: "CASE WHEN{custrecord_efx_ms_con_contact} = 'T' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END WHEN{custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE{custrecord_efx_ms_con_ship_cus.billstate} END END"
                            });
                            var matchState = matchStateCode(state);
                            if (matchState != null) {
                                state = matchState;
                            }
                            var custrecord_efx_ms_con_period = period;
                            var w_ssi = (weight.length !== 0) ? weight : result.getValue({ name: 'custitem_efx_ms_service_weight', join: 'custrecord_efx_ms_con_ssi' });
                            var peso = (parseFloat(custrecord_efx_ms_con_qty) * parseFloat(custrecord_efx_ms_con_w));
                            if (weight.length !== 0) {
                                for (var i = 0; i < listaSello.length; i++) {
                                    if (peso >= parseFloat(listaSello[i].pesoMin) && peso <= parseFloat(listaSello[i].pesoMax) && (listaSello[i].articulo).indexOf(itemNameSSI) !== -1 && (custrecord_efx_ms_con_postmark === listaSello[i].selloPostal)) {
                                        custrecord_efx_ms_con_sm = listaSello[i].methodSent;
                                        paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                                        custrecord_efx_ms_con_postmark = listaSello[i].selloPostal;
                                        break;
                                    }
                                }
                            }
                            var custrecord_efx_ms_con_act = result.getValue({ name: 'custrecord_efx_ms_con_act' });
                            var lastmodified = result.getValue({ name: 'lastmodified' });
                            var entityid = result.getValue({ name: 'entityid', join: 'owner' });
                            var custrecord_efx_ms_con = result.getValue({ name: 'custrecord_efx_ms_con' });
                            var formula_nombre = result.getValue({
                                name: "formulatext_nombre",
                                formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'F' THEN {custrecord_efx_ms_con_ship_cus} || ' ' || {custrecord_efx_ms_con_ship_cus.companyname} WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'T' THEN {custrecord_efx_ms_con_ship_cus} || ' ' || {custrecord_efx_ms_con_ship_cus.firstname} || ' ' || {custrecord_efx_ms_con_ship_cus.lastname} END WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_atention} END"
                            });
                            detailData.push({
                                id: internal_id,
                                custrecord_efx_ms_con_postmark: custrecord_efx_ms_con_postmark,
                                custrecord_efx_ms_con_sd: custrecord_efx_ms_con_sd,
                                custrecord_efx_ms_con_ed: custrecord_efx_ms_con_ed,
                                custrecord_efx_ms_con_contact: custrecord_efx_ms_con_contact,
                                custrecord_efx_ms_con_ship_cont: custrecord_efx_ms_con_ship_cont,
                                custrecord_efx_ms_con_ship_cus: custrecord_efx_ms_con_ship_cus,
                                custrecord_efx_ms_con_addr: custrecord_efx_ms_con_addr,
                                custrecord_efx_ms_con_cust_addr: custrecord_efx_ms_con_cust_addr,
                                custrecord_efx_ms_con_des: custrecord_efx_ms_con_des,
                                custrecord_efx_ms_con_qty: parseInt(custrecord_efx_ms_con_qty),
                                custrecord_efx_ms_con_w: custrecord_efx_ms_con_w,
                                custrecord_efx_ms_con_unit_w: custrecord_efx_ms_con_unit_w,
                                custrecord_efx_ms_con_sm: custrecord_efx_ms_con_sm,
                                custrecord_efx_ms_con_ssi: custrecord_efx_ms_con_ssi,
                                custrecord_efx_ms_con_ssi_w: custrecord_efx_ms_con_ssi_w,
                                formulatext_month: formulatext_month,
                                prefix: prefix,
                                cpCustom: cpCustom,
                                state: state,
                                custrecord_efx_ms_con: custrecord_efx_ms_con,
                                custrecord_efx_ms_con_period: period,
                                custrecord_efx_ms_con_tp: custrecord_efx_ms_con_tp,
                                consecutive: consecutive.toString(),
                                yearApply: fYear.toString(),
                                itemSSI: itemSSI,
                                paqueteria: paqueteria,
                                methodSent: methodSent,
                                peso: peso,
                                con_ship_cus: con_ship_cus,
                                itemNameSSI: itemNameSSI,
                                w_ssi: w_ssi,
                                custrecord_efx_ms_con_act: custrecord_efx_ms_con_act,
                                lastmodified: lastmodified,
                                entityid: entityid,
                                dateRol: dateRol,
                                custrecord_efx_ms_con: custrecord_efx_ms_con,
                                formula_nombre: formula_nombre
                            })
                        }
                    });
                }
                var itemsProcess = getItemSales(itemDataObj, weight);
                if (itemsProcess) {
                    var kyIP = Object.keys(itemsProcess);
                    for (var ip in kyIP) {
                        for (var a in detailData) {
                            for (var b in itemsProcess[kyIP[ip]]) {
                                if (detailData[a].itemSSI === kyIP[ip]) {
                                    if (detailData[a].yearApply === itemsProcess[kyIP[ip]][b].year) {
                                        if (detailData[a].consecutive === itemsProcess[kyIP[ip]][b].consecutive) {
                                            detailData[a].custrecord_efx_ms_con_ssi = itemsProcess[kyIP[ip]][b].nameInventory;
                                            detailData[a].custrecord_efx_ms_con_ssi_w = itemsProcess[kyIP[ip]][b].weight;
                                            detailData[a].formulatext_month = itemsProcess[kyIP[ip]][b].period;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                var customerProcess = getCustomerExtNumber(customersObj);
                if (customerProcess) {
                    var keyCustomer = Object.keys(customerProcess);
                    for (var cust in keyCustomer) {
                        for (var a in detailData) {
                            if (detailData[a].custrecord_efx_ms_con_contact === false) {
                                if (detailData[a].con_ship_cus === keyCustomer[cust]) {
                                    detailData[a].custrecord_efx_ms_con_cust_addr = detailData[a].custrecord_efx_ms_con_cust_addr.replace(' _ ', customerProcess[keyCustomer[cust]].streetnum + ' ');
                                }
                            }
                        }
                    }
                }
                return detailData;
            } catch (e) {
                log.error('Error on getAllDeailContract', e);
            }
        }

        function matchStateCode(state) {
            var states = {
                'AGS': 'Aguascalientes',
                'BC': 'Baja California',
                'BCS': 'Baja California Sur',
                'CAM': 'Campeche',
                'CHIS': 'Chiapas',
                'CHIH': 'Chihuahua',
                'COAH': 'Coahuila',
                'COL': 'Colima',
                'DGO': 'Durango',
                'GTO': 'Guanajuato',
                'GRO': 'Guerrero',
                'HGO': 'Hidalgo',
                'JAL': 'Jalisco',
                'CDMX': 'CDMX',
                'MICH': 'Michoacán',
                'MOR': 'Morelos',
                'MEX': 'Estado de México',
                'NAY': 'Nayarit',
                'NL': 'Nuevo León',
                'OAX': 'Oaxaca',
                'PUE': 'Puebla',
                'QRO': 'Querétaro',
                'QROO': 'Quintana Roo',
                'SLP': 'San Luis Potosí',
                'SIN': 'Sinaloa',
                'SON': 'Sonora',
                'TAB': 'Tabasco',
                'TAMPS': 'Tamaulipas',
                'TLAX': 'Tlaxcala',
                'VER': 'Veracruz',
                'YUC': 'Yucatán',
                'ZAC': 'Zacatecas'
            };
            try {
                if (state != "") {
                    if (states.hasOwnProperty(state)) {
                        return states[state].toUpperCase();
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            } catch (e) {
                log.error({ title: 'Error on matchStateCode', details: e });
                return null;
            }
        }

        function setSortData(dataArray) {
            try {
                dataArray.sort(function (a, b) {
                    if (a.custrecord_efx_ms_con_tp > b.custrecord_efx_ms_con_tp) {
                        return 1;
                    }
                    if (a.custrecord_efx_ms_con_tp < b.custrecord_efx_ms_con_tp) {
                        return -1;
                    }
                    return 0;
                });
                var newArraySort = {};
                for (var i in dataArray) {
                    if (newArraySort.hasOwnProperty(dataArray[i].custrecord_efx_ms_con_tp) === false) {
                        newArraySort[dataArray[i].custrecord_efx_ms_con_tp] = [];
                        newArraySort[dataArray[i].custrecord_efx_ms_con_tp].push(dataArray[i]);
                    } else {
                        newArraySort[dataArray[i].custrecord_efx_ms_con_tp].push(dataArray[i]);
                    }
                }
                var extranjero = {};
                if (newArraySort.hasOwnProperty('Suscripción extranjero') === true) {
                    var dataExtran = newArraySort['Suscripción extranjero'];
                    dataExtran.sort(function (a, b) {
                        if (a.custrecord_efx_ms_con_sm > b.custrecord_efx_ms_con_sm) {
                            return 1;
                        }
                        if (a.custrecord_efx_ms_con_sm < b.custrecord_efx_ms_con_sm) {
                            return -1;
                        }
                        return 0;
                    });
                    for (var i in dataExtran) {
                        if (dataExtran[i].custrecord_efx_ms_con_sm.toUpperCase() === "SEPOMEX  AEREO" || dataExtran[i].custrecord_efx_ms_con_sm.toUpperCase() === "SEPOMEX AEREO CERTIFCIADO") {
                            if (extranjero.hasOwnProperty(dataExtran[i].custrecord_efx_ms_con_sm)) {
                                extranjero[dataExtran[i].custrecord_efx_ms_con_sm].push(dataExtran[i]);
                            } else {
                                extranjero[dataExtran[i].custrecord_efx_ms_con_sm] = [];
                                extranjero[dataExtran[i].custrecord_efx_ms_con_sm].push(dataExtran[i]);
                            }
                        }
                    }
                    var keysExtranjero = Object.keys(extranjero);
                    var objTemp = [];
                    for (var i in keysExtranjero) {
                        extranjero[keysExtranjero[i]].sort(function (a, b) {
                            if (a.custrecord_efx_ms_con_qty > b.custrecord_efx_ms_con_qty) {
                                return 1;
                            }
                            if (a.custrecord_efx_ms_con_qty < b.custrecord_efx_ms_con_qty) {
                                return -1;
                            }
                            return 0;
                        });
                    }
                    for (var key in keysExtranjero) {
                        for (var i in extranjero[keysExtranjero[key]]) {
                            objTemp.push(extranjero[keysExtranjero[key]][i])
                        }
                    }
                    newArraySort['Suscripción extranjero'] = objTemp;
                }
                var nacional = {};
                if (newArraySort.hasOwnProperty('Suscripción nacional') === true) {
                    var nacData = newArraySort['Suscripción nacional'];
                    nacData.sort(function (a, b) {
                        if (a.custrecord_efx_ms_con_sm > b.custrecord_efx_ms_con_sm) {
                            return 1;
                        }
                        if (a.custrecord_efx_ms_con_sm < b.custrecord_efx_ms_con_sm) {
                            return -1;
                        }
                        return 0;
                    });
                    for (var i in nacData) {
                        if (nacional.hasOwnProperty(nacData[i].custrecord_efx_ms_con_sm)) {
                            nacional[nacData[i].custrecord_efx_ms_con_sm].push(nacData[i]);
                        } else {
                            nacional[nacData[i].custrecord_efx_ms_con_sm] = [];
                            nacional[nacData[i].custrecord_efx_ms_con_sm].push(nacData[i]);
                        }
                    }
                    var keyNac = Object.keys(nacional);
                    var nacionalFilter = {};
                    for (var key in keyNac) {
                        if (keyNac[key] === 'SEPOMEX ORDINARIO') {
                            nacionalFilter[keyNac[key]] = nacional[keyNac[key]]
                        }
                        if (keyNac[key] === 'SEPOMEX CERTIFICADO') {
                            nacionalFilter[keyNac[key]] = nacional[keyNac[key]]
                        }

                        if (keyNac[key].indexOf('LIBRERÍA') === 0) {
                            if (nacionalFilter.hasOwnProperty('Librerías') === true) {
                                nacionalFilter['Librerías'][keyNac[key]] = nacional[keyNac[key]];
                            } else {
                                nacionalFilter['Librerías'] = {};
                                nacionalFilter['Librerías'][keyNac[key]] = nacional[keyNac[key]];
                            }
                        }
                        if (keyNac[key] === 'SEPOMEX PAQ-POST') {
                            nacionalFilter[keyNac[key]] = nacional[keyNac[key]]
                        }
                        if (keyNac[key] === 'DHL NORMAL NACIONAL') {
                            nacionalFilter[keyNac[key]] = nacional[keyNac[key]]
                        }
                        if (keyNac[key] === 'FEDEX NACIONAL') {
                            nacionalFilter[keyNac[key]] = nacional[keyNac[key]]
                        }
                        if (keyNac[key] === 'REDPACK NACIONAL') {
                            nacionalFilter[keyNac[key]] = nacional[keyNac[key]]
                        }


                        if (keyNac[key] != 'SEPOMEX ORDINARIO' &&
                            keyNac[key] != 'SEPOMEX CERTIFICADO' &&
                            keyNac[key].indexOf('LIBRERÍA') < 0 &&
                            keyNac[key] != 'SEPOMEX PAQ-POST' &&
                            keyNac[key] != 'DHL NORMAL NACIONAL' &&
                            keyNac[key] != 'FEDEX NACIONAL' &&
                            keyNac[key] != 'REDPACK NACIONAL') {
                            if (nacionalFilter.hasOwnProperty('Sin Logica') === true) {
                                nacionalFilter['Sin Logica'][keyNac[key]] = nacional[keyNac[key]];
                            } else {
                                nacionalFilter['Sin Logica'] = {};
                                nacionalFilter['Sin Logica'][keyNac[key]] = nacional[keyNac[key]];
                            }
                        }
                    }
                    if (nacionalFilter.hasOwnProperty('Librerías')) {
                        var librerias = nacionalFilter['Librerías'];
                        var keyLib = Object.keys(librerias);
                        for (var i in keyLib) {
                            librerias[keyLib[i]].sort(function (a, b) {
                                if (a.custrecord_efx_ms_con_qty > b.custrecord_efx_ms_con_qty) {
                                    return 1;
                                }
                                if (a.custrecord_efx_ms_con_qty < b.custrecord_efx_ms_con_qty) {
                                    return -1;
                                }
                                return 0;
                            });
                        }
                        var temp = [];
                        for (var i in keyLib) {
                            for (var libreriaKey in librerias[keyLib[i]]) {
                                temp.push(librerias[keyLib[i]][libreriaKey]);
                            }
                        }
                        nacionalFilter['Librerías'] = temp;
                    }

                    if (nacionalFilter.hasOwnProperty('SEPOMEX ORDINARIO')) {
                        var ordinarioData = nacionalFilter['SEPOMEX ORDINARIO'];
                        ordinarioData = orderData(ordinarioData, 'custrecord_efx_ms_con_qty');
                        ordinarioData = groupData(ordinarioData, 'custrecord_efx_ms_con_qty', false);
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(ordinarioData[keysordinarioqty[k]], 'state');
                            ordinarioData[keysordinarioqty[k]] = groupData(ordinarioData[keysordinarioqty[k]], 'state', true)
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] = orderData(ordinarioData[keysordinarioqty[k]][keysState[j]], 'cpCustom')
                            }
                        }
                        var temp = [];
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                var stateLocal = ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['SEPOMEX ORDINARIO'] = temp;
                    }

                    if (nacionalFilter.hasOwnProperty('SEPOMEX CERTIFICADO')) {
                        var ordinarioData = nacionalFilter['SEPOMEX CERTIFICADO'];
                        ordinarioData = orderData(ordinarioData, 'custrecord_efx_ms_con_qty');
                        ordinarioData = groupData(ordinarioData, 'custrecord_efx_ms_con_qty', false);
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(ordinarioData[keysordinarioqty[k]], 'state');
                            ordinarioData[keysordinarioqty[k]] = groupData(ordinarioData[keysordinarioqty[k]], 'state', true)
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] = orderData(ordinarioData[keysordinarioqty[k]][keysState[j]], 'cpCustom')
                            }
                        }
                        var temp = [];
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                var stateLocal = ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['SEPOMEX CERTIFICADO'] = temp;
                    }

                    if (nacionalFilter.hasOwnProperty('SEPOMEX PAQ-POST')) {
                        var ordinarioData = nacionalFilter['SEPOMEX PAQ-POST'];
                        ordinarioData = orderData(ordinarioData, 'custrecord_efx_ms_con_qty');
                        ordinarioData = groupData(ordinarioData, 'custrecord_efx_ms_con_qty', false);
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(ordinarioData[keysordinarioqty[k]], 'state');
                            ordinarioData[keysordinarioqty[k]] = groupData(ordinarioData[keysordinarioqty[k]], 'state', true)
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] = orderData(ordinarioData[keysordinarioqty[k]][keysState[j]], 'cpCustom')
                            }
                        }
                        var temp = [];
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                var stateLocal = ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['SEPOMEX PAQ-POST'] = temp;
                    }

                    if (nacionalFilter.hasOwnProperty('DHL NORMAL NACIONAL')) {
                        var ordinarioData = nacionalFilter['DHL NORMAL NACIONAL'];
                        ordinarioData = orderData(ordinarioData, 'custrecord_efx_ms_con_qty');
                        ordinarioData = groupData(ordinarioData, 'custrecord_efx_ms_con_qty', false);
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(ordinarioData[keysordinarioqty[k]], 'state');
                            ordinarioData[keysordinarioqty[k]] = groupData(ordinarioData[keysordinarioqty[k]], 'state', true)
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] = orderData(ordinarioData[keysordinarioqty[k]][keysState[j]], 'cpCustom')
                            }
                        }
                        var temp = [];
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                var stateLocal = ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['DHL NORMAL NACIONAL'] = temp;
                    }

                    if (nacionalFilter.hasOwnProperty('FEDEX NACIONAL')) {
                        var ordinarioData = nacionalFilter['FEDEX NACIONAL'];
                        ordinarioData = orderData(ordinarioData, 'custrecord_efx_ms_con_qty');
                        ordinarioData = groupData(ordinarioData, 'custrecord_efx_ms_con_qty', false);
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(ordinarioData[keysordinarioqty[k]], 'state');
                            ordinarioData[keysordinarioqty[k]] = groupData(ordinarioData[keysordinarioqty[k]], 'state', true)
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] = orderData(ordinarioData[keysordinarioqty[k]][keysState[j]], 'cpCustom')
                            }
                        }
                        var temp = [];
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                var stateLocal = ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['FEDEX NACIONAL'] = temp;
                    }

                    if (nacionalFilter.hasOwnProperty('REDPACK NACIONAL')) {
                        var ordinarioData = nacionalFilter['REDPACK NACIONAL'];
                        ordinarioData = orderData(ordinarioData, 'custrecord_efx_ms_con_qty');
                        ordinarioData = groupData(ordinarioData, 'custrecord_efx_ms_con_qty', false);
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(ordinarioData[keysordinarioqty[k]], 'state');
                            ordinarioData[keysordinarioqty[k]] = groupData(ordinarioData[keysordinarioqty[k]], 'state', true)
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] = orderData(ordinarioData[keysordinarioqty[k]][keysState[j]], 'cpCustom')
                            }
                        }
                        var temp = [];
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(ordinarioData[keysordinarioqty[k]]);
                            for (var j in keysState) {
                                var stateLocal = ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['REDPACK NACIONAL'] = temp;
                    }

                    newArraySort['Suscripción nacional'] = nacionalFilter;

                }
                // "posisiones": ["Extranjero", "SEPOMEX ORDINARIO", "SEPOMEX CERTIFICADO", "LIBRERIAS", "SEPOMEX PAQ-POST", "DHL NORMAL NACIONAL", "FEDEX NACIONAL", "REDPACK NACIONAL"]
                var finalObj = [];
                if (newArraySort.hasOwnProperty('Suscripción extranjero')) {
                    for (var i in newArraySort['Suscripción extranjero']) {
                        finalObj.push(newArraySort['Suscripción extranjero'][i]);
                    }
                }
                if (newArraySort['Suscripción nacional'].hasOwnProperty('SEPOMEX ORDINARIO')) {
                    for (var i in newArraySort['Suscripción nacional']['SEPOMEX ORDINARIO']) {
                        finalObj.push(newArraySort['Suscripción nacional']['SEPOMEX ORDINARIO'][i]);
                    }
                }
                if (newArraySort['Suscripción nacional'].hasOwnProperty('SEPOMEX CERTIFICADO')) {
                    for (var i in newArraySort['Suscripción nacional']['SEPOMEX CERTIFICADO']) {
                        finalObj.push(newArraySort['Suscripción nacional']['SEPOMEX CERTIFICADO'][i]);
                    }
                }
                if (newArraySort['Suscripción nacional'].hasOwnProperty('Librerías')) {
                    for (var i in newArraySort['Suscripción nacional']['Librerías']) {
                        finalObj.push(newArraySort['Suscripción nacional']['Librerías'][i]);
                    }
                }
                if (newArraySort['Suscripción nacional'].hasOwnProperty('SEPOMEX PAQ-POST')) {
                    for (var i in newArraySort['Suscripción nacional']['SEPOMEX PAQ-POST']) {
                        finalObj.push(newArraySort['Suscripción nacional']['SEPOMEX PAQ-POST'][i]);
                    }
                }
                if (newArraySort['Suscripción nacional'].hasOwnProperty('DHL NORMAL NACIONAL')) {
                    for (var i in newArraySort['Suscripción nacional']['DHL NORMAL NACIONAL']) {
                        finalObj.push(newArraySort['Suscripción nacional']['DHL NORMAL NACIONAL'][i]);
                    }
                }
                if (newArraySort['Suscripción nacional'].hasOwnProperty('FEDEX NACIONAL')) {
                    for (var i in newArraySort['Suscripción nacional']['FEDEX NACIONAL']) {
                        finalObj.push(newArraySort['Suscripción nacional']['FEDEX NACIONAL'][i]);
                    }
                }
                if (newArraySort['Suscripción nacional'].hasOwnProperty('REDPACK NACIONAL')) {
                    for (var i in newArraySort['Suscripción nacional']['REDPACK NACIONAL']) {
                        finalObj.push(newArraySort['Suscripción nacional']['REDPACK NACIONAL'][i]);
                    }
                }

                return finalObj;
            } catch (e) {
                log.error({ title: 'Error on setSortData', details: e });
                return {}
            }
        }

        function orderData(data, key) {
            data.sort(function (a, b) {
                if (a[key] > b[key]) {
                    return 1;
                }
                if (a[key] < b[key]) {
                    return -1;
                }
                return 0;
            });
            return data;
        }

        function groupData(data, key, str) {
            var obj = {}
            if (str === false) {
                for (var dataKey in data) {
                    if (obj.hasOwnProperty(data[dataKey][key])) {
                        obj[data[dataKey][key]].push(data[dataKey])
                    } else {
                        obj[data[dataKey][key]] = [];
                        obj[data[dataKey][key]].push(data[dataKey]);
                    }
                }
            } else {
                for (var dataKey in data) {
                    if (obj.hasOwnProperty(data[dataKey][key].toUpperCase())) {
                        obj[data[dataKey][key].toUpperCase()].push(data[dataKey])
                    } else {
                        obj[data[dataKey][key].toUpperCase()] = [];
                        obj[data[dataKey][key].toUpperCase()].push(data[dataKey]);
                    }
                }
            }
            return obj;
        }

        return {
            onRequest: onRequest
        };

    });
