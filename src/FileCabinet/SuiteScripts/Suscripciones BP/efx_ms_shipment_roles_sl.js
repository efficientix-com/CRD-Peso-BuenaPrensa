/**
 * @NApiVersion 2.x
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 */
define(['N/record', 'N/search', 'N/ui/serverWidget', 'N/log', 'N/format', 'N/file', 'N/redirect', 'N/runtime', 'N/url', 'N/task'],
    /**
     * @param{record} record
     * @param{search} search
     * @param{serverWidget} serverWidget
     * @param{log} log
     * @param{format} format
     * @param{file} file
     * @param{redirect} redirect
     * @param{runtime} runtime
     * @param{url} url
     * @param{task} task
     */
    function (record, search, serverWidget, log, format, file, redirect, runtime, url, task) {

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
                var pageIndex = (params.pagein) ? params.pagein : 0;
                var itemFilter = params.itemfilter;
                var periodFilter = params.periodfilter;
                var yearFilter = (params.yearFilter) ? params.yearFilter : new Date().getFullYear();
                log.audit("📍 ~ onRequest ~ yearFilter", yearFilter);
                log.audit("📍 ~ onRequest ~ itemFilter", itemFilter);
                if (itemFilter) {
                    peso = getPesoForSuscripcion(itemFilter);
                }
                var form = createInterface();
                log.audit({
                    title: 'Gobernance use', details: {
                        detail: 'After gobernance of searchrol() and createInterface()',
                        use: runtime.getCurrentScript().getRemainingUsage()
                    }
                });
                if (params.fileID) {
                    var fieldFileUrl = form.getField({
                        id: 'custpage_file_url'
                    });
                    fieldFileUrl.defaultValue = getFileURL(params.fileID);
                }
                if (params.errorMsg) {
                    var fieldMsgError = form.getField({
                        id: 'custpage_msg_error'
                    });
                    fieldMsgError.defaultValue = params.errorMsg;
                }
                if (yearFilter) {
                    var fieldYearFilter = form.getField({
                        id: 'custpage_year_apply'
                    });
                    fieldYearFilter.defaultValue = yearFilter.toString();
                }

                if (itemFilter && periodFilter && yearFilter) {
                    var rolDetail = getSearchRols(periodFilter);
                    var searchObj = searchContractDetails(rolDetail, itemFilter, periodFilter, yearFilter);
                    var listPages = getListPage(searchObj);
                    log.audit({
                        title: 'Gobernance use', details: {
                            detail: 'After gobernance of searchContractDetails() and listPage()',
                            use: runtime.getCurrentScript().getRemainingUsage()
                        }
                    });
                    var detailsResult = [];


                    if (params.itemfilter) {
                        var fieldItemFilter = form.getField({
                            id: 'custpage_item_filter'
                        });
                        fieldItemFilter.defaultValue = itemFilter;
                        var fieldPeriodFilter = form.getField({
                            id: 'custpage_period_filter'
                        });
                        fieldPeriodFilter.defaultValue = periodFilter;

                    }
                    if (listPages.length) {

                        form.addButton({
                            id: 'button_export',
                            label: 'Exportar',
                            functionName: 'exportData'
                        });

                        if (params.export) {
                            var searchAllObj = searchAllContracts(rolDetail, itemFilter, periodFilter, yearFilter.toString());
                            log.audit({
                                title: 'Gobernance use', details: {
                                    detail: 'After gobernance of searchAllContracts()',
                                    use: runtime.getCurrentScript().getRemainingUsage()
                                }
                            });
                            if (searchAllObj.countError === 0) {
                                // response.writeLine({output: JSON.stringify(searchAllObj.dataRol)})
                                var rolID = createRol(periodFilter, peso);
                                // Obtención del Header para el CSV
                                var csvHeader = 'custrecord_efx_ms_dre_re,custrecord_efx_ms_dre_item,custrecord_efx_ms_dre_sm,custrecord_efx_ms_dre_detail_contract';
                                if (rolID) {
                                    if (searchAllObj.dataRol.length) {
                                        var detailForRol = searchAllObj.dataRol
                                        // Iteración de lineas a agregar al layout
                                        var fileContent = csvHeader + "\n";
                                        // Lineas del archivo
                                        for (var i = 0; i < detailForRol.length; i++) {
                                            var line = rolID + ',' + detailForRol[i].custrecord_efx_ms_dre_item + ',' + detailForRol[i].custrecord_efx_ms_dre_sm + ',' + detailForRol[i].custrecord_efx_ms_dre_detail_contract;
                                            fileContent += line + "\n";
                                        }
                                        // Guardado del archivo
                                        var fileObj = file.create({
                                            name: 'detailForRol.csv',
                                            fileType: file.Type.CSV,
                                            contents: fileContent,
                                            encoding: file.Encoding.UTF8,
                                            folder: -15,
                                        });
                                        var childFileId = fileObj.save();
                                        var savedCSVImportId = runtime.getCurrentScript().getParameter({
                                            name: 'custscript_efx_ms_dre_import_csv'
                                        });
                                        log.audit({
                                            title: 'data import', details: {
                                                1: childFileId,
                                                2: savedCSVImportId,
                                                3: (childFileId && savedCSVImportId)
                                            }
                                        });
                                        if (childFileId && savedCSVImportId) {
                                            var scheduledTask = task.create({
                                                taskType: task.TaskType.SCHEDULED_SCRIPT
                                            });
                                            scheduledTask.scriptId = 'customscript_efx_ms_shipment_roles_ss';
                                            scheduledTask.deploymentId = 'customdeploy_efx_ms_shipment_roles_ss';
                                            scheduledTask.params = {
                                                custscript_rol_id: rolID,
                                                custscript_csv_import: savedCSVImportId,
                                                custscript_file_id: childFileId
                                            }
                                            var scheduledSubmit = scheduledTask.submit();
                                            log.audit({ title: 'scheduledSubmit', details: scheduledSubmit });
                                            if (scheduledSubmit) {
                                                redirect.toRecord({
                                                    type: 'customrecord_efx_ms_rol_env',
                                                    id: rolID,
                                                    isEditMode: false,
                                                    parameters: {}
                                                });
                                            }
                                        }
                                    }
                                }
                            } else {
                                log.error({ title: 'CountData Error', details: { count: searchAllObj.countError, details: searchAllObj.rolErrors } });
                                redirect.toSuitelet({
                                    scriptId: 'customscript_efx_ms_shipment_roles_sl',
                                    deploymentId: 'customdeploy_efx_ms_shipment_roles_sl',
                                    parameters: { 'errorMsg': 'T' }
                                });
                            }
                        }

                        detailsResult = getPageData(searchObj, pageIndex);
                        var fieldPages = form.getField({
                            id: 'custpage_pages'
                        });
                        for (var i = 0; i < listPages.length; i++) {
                            var page = listPages[i];
                            fieldPages.addSelectOption({
                                value: page.index,
                                text: page.text
                            });
                        }
                        fieldPages.defaultValue = pageIndex;

                        var sublist = form.getSublist({ id: 'sublist_roles_genered' });
                        for (var j = 0; j < detailsResult.length; j++) {
                            var output = url.resolveRecord({
                                recordType: 'customrecord_efx_ms_contract_detail',
                                recordId: detailsResult[j].id,
                                isEditMode: false
                            });
                            sublist.setSublistValue({
                                id: 'sublist_id_detail',
                                line: j,
                                value: "<a href=" + output + ">" + detailsResult[j].id + "</a>"

                            });
                            sublist.setSublistValue({
                                id: 'sublist_item',
                                line: j,
                                value: detailsResult[j].item
                            });
                            sublist.setSublistValue({
                                id: 'sublist_quantity',
                                line: j,
                                value: detailsResult[j].quantity
                            });
                            sublist.setSublistValue({
                                id: 'sublist_weight',
                                line: j,
                                value: detailsResult[j].weight
                            });
                            sublist.setSublistValue({
                                id: 'sublist_third',
                                line: j,
                                value: (detailsResult[j].third) ? 'T' : 'F'
                            });
                            sublist.setSublistValue({
                                id: 'sublist_address',
                                line: j,
                                value: (detailsResult[j].address) ? detailsResult[j].address : ' '
                            });
                            if (detailsResult[j].shipmethod) {
                                sublist.setSublistValue({
                                    id: 'sublist_parcel',
                                    line: j,
                                    value: detailsResult[j].shipmethod
                                });
                            }
                            var dateCreated = detailsResult[j].datecreated.split(' ');
                            sublist.setSublistValue({
                                id: 'sublist_date_created',
                                line: j,
                                value: dateCreated[0]
                            });
                            sublist.setSublistValue({
                                id: 'sublist_created_by',
                                line: j,
                                value: detailsResult[j].owner
                            });
                            var dateUpdated = detailsResult[j].dateupdate.split(' ');
                            sublist.setSublistValue({
                                id: 'sublist_date_updated',
                                line: j,
                                value: dateUpdated[0]
                            });
                        }
                    }
                }
                response.writePage(form);
            } catch (e) {
                log.error('Error on onRequest', e)
                redirect.toSuitelet({
                    scriptId: 'customscript_efx_ms_shipment_roles_sl',
                    deploymentId: 'customdeploy_efx_ms_shipment_roles_sl',
                    parameters: { errorMsg: 'T' }
                });
            }
        }
        function getPesoForSuscripcion(id) {
            try {
                var idSuscripcion = '';
                var customrecord_efx_ms_sus_relatedSearchObj = search.create({
                    type: "customrecord_efx_ms_sus_related",
                    filters:
                        [
                            ["custrecord_efx_ms_inventory_item", "anyof", id]
                        ],
                    columns:
                        [
                            search.createColumn({ name: "custrecord_efx_ms_item_sus", label: "Artículo de Suscripción" }),
                            search.createColumn({ name: "internalid", join: "CUSTRECORD_EFX_MS_ITEM_SUS", label: "Internal ID" })
                        ]
                });
                var searchResultCount = customrecord_efx_ms_sus_relatedSearchObj.runPaged().count;
                log.debug("customrecord_efx_ms_sus_relatedSearchObj result count", searchResultCount);
                customrecord_efx_ms_sus_relatedSearchObj.run().each(function (result) {
                    // .run().each has a limit of 4,000 results
                    idSuscripcion = result.getValue({ name: "internalid", join: "CUSTRECORD_EFX_MS_ITEM_SUS" });
                    log.audit({title: 'idSuscripcion', details: idSuscripcion});
                    return true;
                });
                var peso ='';
                if(idSuscripcion.length === 0 ){
                    peso = search.lookupFields({ type: search.Type.ITEM, id: id, columns: ['custitem_efx_ms_service_weight'] });
                }else{
                    peso = search.lookupFields({ type: search.Type.ITEM, id: idSuscripcion, columns: ['custitem_efx_ms_service_weight'] });
                }
                log.audit({ title: 'peso', details: peso });
                return peso.custitem_efx_ms_service_weight;

            } catch (error) {
                log.error({ title: 'Error getPesoForSuscripcion', details: error });
                return '';
            }
        }
        function createInterface() {
            try {

                var period = [
                    { text: 'Enero', value: 1 },
                    { text: 'Febrero', value: 2 },
                    { text: 'Marzo', value: 3 },
                    { text: 'Abril', value: 4 },
                    { text: 'Mayo', value: 5 },
                    { text: 'Junio', value: 6 },
                    { text: 'Julio', value: 7 },
                    { text: 'Agosto', value: 8 },
                    { text: 'Septiembre', value: 9 },
                    { text: 'Octubre', value: 10 },
                    { text: 'Noviembre', value: 11 },
                    { text: 'Diciembre', value: 12 },
                ];

                var form = serverWidget.createForm({
                    title: 'MS - Generación de roles SL'
                });

                form.clientScriptModulePath = './efx_ms_shipment_roles_cs.js';

                form.addButton({
                    id: 'button_filter',
                    label: 'Filtrar',
                    functionName: 'filterData'
                });

                var fieldgroup = form.addFieldGroup({
                    id: 'fieldgroup_general',
                    label: 'Criterios de búsqueda'
                });

                var fieldFileUrl = form.addField({
                    id: 'custpage_file_url',
                    type: serverWidget.FieldType.TEXTAREA,
                    label: 'URL File',
                    container: 'fieldgroup_general'
                });
                fieldFileUrl.updateDisplayType({ displayType: serverWidget.FieldDisplayType.HIDDEN });
                var fieldMsgError = form.addField({
                    id: 'custpage_msg_error',
                    type: serverWidget.FieldType.TEXT,
                    label: 'Error Message',
                    container: 'fieldgroup_general'
                });
                fieldMsgError.updateDisplayType({ displayType: serverWidget.FieldDisplayType.HIDDEN });
                var fieldItemFilter = form.addField({
                    id: 'custpage_item_filter',
                    type: serverWidget.FieldType.SELECT,
                    label: 'Artículo',
                    container: 'fieldgroup_general',
                    source: 'item'
                });
                var fieldPeriod = form.addField({
                    id: 'custpage_period_filter',
                    type: serverWidget.FieldType.SELECT,
                    label: 'Periodo',
                    container: 'fieldgroup_general'
                });
                fieldPeriod.addSelectOption({
                    value: '',
                    text: ''
                });
                for (var i = 0; i < period.length; i++) {
                    fieldPeriod.addSelectOption({
                        value: period[i].value,
                        text: period[i].text
                    });
                }
                var fieldPages = form.addField({
                    id: 'custpage_pages',
                    type: serverWidget.FieldType.SELECT,
                    label: 'Pagina',
                    container: 'fieldgroup_general'
                });
                var fieldYearly = form.addField({
                    id: 'custpage_year_apply',
                    type: serverWidget.FieldType.TEXT,
                    label: 'Año en que aplica',
                    container: 'fieldgroup_general'
                });

                var sublist = form.addSublist({
                    id: 'sublist_roles_genered',
                    type: serverWidget.SublistType.LIST,
                    label: 'Roles'
                });

                var subFieldDetailID = sublist.addField({
                    id: 'sublist_id_detail',
                    type: serverWidget.FieldType.TEXT,
                    label: 'ID'
                });
                var subFieldItem = sublist.addField({
                    id: 'sublist_item',
                    type: serverWidget.FieldType.SELECT,
                    label: 'Artículo',
                    source: 'item'
                });
                var subFieldQuantity = sublist.addField({
                    id: 'sublist_quantity',
                    type: serverWidget.FieldType.TEXT,
                    label: 'Cantidad'
                });
                var subFieldWeight = sublist.addField({
                    id: 'sublist_weight',
                    type: serverWidget.FieldType.TEXT,
                    label: 'Peso'
                });
                var subFieldThird = sublist.addField({
                    id: 'sublist_third',
                    type: serverWidget.FieldType.CHECKBOX,
                    label: 'Tercero'
                });
                var subFieldAddress = sublist.addField({
                    id: 'sublist_address',
                    type: serverWidget.FieldType.TEXTAREA,
                    label: 'Dirección'
                });
                var subFieldShipMet = sublist.addField({
                    id: 'sublist_parcel',
                    type: serverWidget.FieldType.TEXT,
                    label: 'Paquetería'
                });
                var subFieldDateCreated = sublist.addField({
                    id: 'sublist_date_created',
                    type: serverWidget.FieldType.DATE,
                    label: 'Fecha de creación'
                });
                var subFieldOnwer = sublist.addField({
                    id: 'sublist_created_by',
                    type: serverWidget.FieldType.TEXT,
                    label: 'Creado por'
                });
                var subFieldDateUpdated = sublist.addField({
                    id: 'sublist_date_updated',
                    type: serverWidget.FieldType.DATE,
                    label: 'Fecha de actualización'
                });

                fieldItemFilter.isMandatory = true;
                fieldPeriod.isMandatory = true;
                fieldYearly.isMandatory = true;

                subFieldDetailID.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldItem.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldQuantity.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldWeight.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldThird.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldAddress.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldShipMet.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldDateCreated.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldOnwer.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });
                subFieldDateUpdated.updateDisplayType({ displayType: serverWidget.FieldDisplayType.INLINE });

                return form;
            } catch (e) {
                log.error('Error on createInterface', e)
            }
        }

        function getSearchRols(period) {
            try {
                var searchObj = search.load({
                    id: 'customsearch_efx_ms_rol_detail'
                });
                if (period) {
                    var periodFilter = search.createFilter({
                        name: 'custrecord_efx_ms_re_period',
                        join: 'custrecord_efx_ms_dre_re',
                        operator: search.Operator.IS,
                        values: period
                    });
                    searchObj.filters.push(periodFilter);
                }
                var idDetails = []; var myPagedResults = searchObj.runPaged({
                    pageSize: 1000
                });
                var thePageRanges = myPagedResults.pageRanges;
                for (var i in thePageRanges) {
                    var thepageData = myPagedResults.fetch({
                        index: thePageRanges[i].index
                    });
                    thepageData.data.forEach(function (result) {
                        idDetails.push(result.getValue({ name: 'custrecord_efx_ms_dre_detail_contract' }));
                    });
                }
                /*searchObj.run().each(function (result) {

                    return true;
                });*/
                return (idDetails.length) ? idDetails : [];
            } catch (e) {
                log.error('Error on getSearchRols', e);
            }
        }

        function searchContractDetails(filterDetail, itemFilter, periodo, yearApply) {
            try {
                var fecha_ini = matchDate(periodo, 'inicio', yearApply)
                var fecha_fin = matchDate(periodo, 'fin', yearApply)
                var conArray = [];
                for (var i = 0; i < filterDetail.length; i++) {
                    conArray.push(filterDetail[i]);
                }
                log.audit({ title: 'Dates', details: { start: fecha_ini, end: fecha_fin } });
                log.audit({ title: 'contratos aplicados', details: conArray });
                log.audit({ title: 'contratos aplicados cantidad', details: conArray.length });
                var filters = [
                    ['internalid', search.Operator.NONEOF, (conArray.length) ? conArray : '@NONE@'],
                    'and',
                    ['custrecord_efx_ms_con_ssi', search.Operator.IS, itemFilter],
                    'and',
                    ['isinactive', search.Operator.IS, 'F'],
                    'and',
                    ['custrecord_efx_ms_con_act', search.Operator.IS, 'T'],
                    'and',
                    /*['custrecord_efx_ms_con.custrecord_efx_ms_con_can', search.Operator.IS, 'F'],
                    'and',
                    ['custrecord_efx_ms_con.isinactive', search.Operator.IS, 'F'],
                    "AND",*/
                    [
                        ["custrecord_efx_ms_con_sd", search.Operator.ONORBEFORE, fecha_ini],
                        "OR",
                        ["custrecord_efx_ms_con_sd", search.Operator.WITHIN, fecha_ini, fecha_fin]
                    ],
                    "AND",
                    [
                        ["custrecord_efx_ms_con_ed", search.Operator.ONORAFTER, fecha_fin],
                        "OR",
                        ["custrecord_efx_ms_con_ed", search.Operator.WITHIN, fecha_ini, fecha_fin]
                    ],
                    /* ["custrecord_efx_ms_con_sd",search.Operator.WHIT,fecha_ini],
                    "AND",
                    ["custrecord_efx_ms_con_ed", search.Operator.ONORAFTER, fecha_fin], */
                    "AND",
                    ["custrecord_efx_ms_sal_ord_related", search.Operator.NONEOF, "@NONE@"],
                    "AND",
                    ["custrecord_efx_ms_con_sales_order_mirror", search.Operator.NONEOF, "@NONE@"]
                ];
                log.audit({ title: 'Filter period', details: { start: fecha_ini, end: fecha_fin } });
                log.audit({ title: 'Filters in search contract details', details: filters });
                var efx_ms_contract_detailSearch = search.create({
                    type: "customrecord_efx_ms_contract_detail",
                    filters: filters,
                    columns:
                        [
                            { name: "internalid" },
                            { name: "custrecord_efx_ms_con_ssi" },
                            { name: "custrecord_efx_ms_con_qty" },
                            { name: "custrecord_efx_ms_con_w" },
                            { name: "custrecord_efx_ms_con_unit_w" },
                            { name: "custrecord_efx_ms_con_contact" },
                            { name: "custrecord_efx_ms_con_addr" },
                            { name: "custrecord_efx_ms_con_sm" },
                            { name: "created" },
                            {
                                name: "entityid",
                                join: "owner"
                            },
                            { name: "lastmodified" }
                        ]
                });
                var searchResultCount = efx_ms_contract_detailSearch.runPaged({ pageSize: 50 });
                return searchResultCount;
            } catch (e) {
                log.error('Error on searchContractDetail', e);
            }
        }

        function getListPage(searchPaged) {
            try {
                var objectResult = [];
                searchPaged.pageRanges.forEach(function (pageRange) {
                    objectResult.push({ index: pageRange.index, text: pageRange.compoundLabel });
                });
                return objectResult;
            } catch (error) {
                log.error({ title: 'getListPage error', details: error });
                return [];
            }
        }

        function getPageData(searchPaged, page) {
            try {
                var pageData = searchPaged.fetch({ index: page });
                var resultData = [];
                pageData.data.forEach(function (result) {
                    var idDetail = result.getValue({ name: "internalid" });
                    var item = result.getValue({ name: "custrecord_efx_ms_con_ssi" });
                    var quantity = result.getValue({ name: "custrecord_efx_ms_con_qty" });
                    var weight = result.getValue({ name: "custrecord_efx_ms_con_w" });
                    var unitWeight = result.getText({ name: "custrecord_efx_ms_con_unit_w" });
                    var third = result.getValue({ name: "custrecord_efx_ms_con_contact" });
                    var address = result.getValue({ name: "custrecord_efx_ms_con_addr" });
                    var shipmethod = result.getText({ name: "custrecord_efx_ms_con_sm" });
                    var datecreated = result.getValue({ name: "created" });
                    var owner = result.getValue({
                        name: "entityid",
                        join: "owner"
                    });
                    var dateupdate = result.getValue({ name: "lastmodified" });
                    resultData.push({
                        id: idDetail,
                        item: item,
                        quantity: quantity,
                        weight: weight + ' ' + unitWeight,
                        third: third,
                        address: address,
                        shipmethod: shipmethod,
                        datecreated: datecreated,
                        owner: owner,
                        dateupdate: dateupdate
                    });
                });
                return resultData;

            } catch (error) {
                log.error({ title: 'getPageData error', details: error });
            }
        }

        function searchAllContracts(filterDetail, itemFilter, periodFilter, yearApply) {
            try {
                var conArray = [];
                for (var i = 0; i < filterDetail.length; i++) {
                    conArray.push(filterDetail[i]);
                }
                var fecha_ini = matchDate(periodFilter, 'inicio', yearApply);
                var fecha_fin = matchDate(periodFilter, 'fin', yearApply);
                var data = [];
                var dataRol = [];
                var rolErrors = [];
                var dateData = new Date();
                var today = periodFilter + '/' + dateData.getDate() + '/' + dateData.getFullYear();
                var filters = [
                    ['internalid', search.Operator.NONEOF, (conArray.length) ? conArray : '@NONE@'],
                    'and',
                    ['custrecord_efx_ms_con_ssi', search.Operator.IS, itemFilter],
                    'and',
                    ['isinactive', search.Operator.IS, 'F'],
                    'and',
                    ['custrecord_efx_ms_con_act', search.Operator.IS, 'T'],
                    "and",
                    [
                        ["custrecord_efx_ms_con_sd", search.Operator.ONORBEFORE, fecha_ini],
                        "OR",
                        ["custrecord_efx_ms_con_sd", search.Operator.WITHIN, fecha_ini, fecha_fin]
                    ],
                    "AND",
                    [
                        ["custrecord_efx_ms_con_ed", search.Operator.ONORAFTER, fecha_fin],
                        "OR",
                        ["custrecord_efx_ms_con_ed", search.Operator.WITHIN, fecha_ini, fecha_fin]
                    ],
                    "AND",
                    ["custrecord_efx_ms_sal_ord_related", search.Operator.NONEOF, "@NONE@"],
                    "AND",
                    ["custrecord_efx_ms_con_sales_order_mirror", search.Operator.NONEOF, "@NONE@"]
                ]
                var efx_ms_contract_detailSearch = search.create({
                    type: "customrecord_efx_ms_contract_detail",
                    filters: filters,
                    columns:
                        [
                            { name: "internalid" },
                            { name: "custrecord_efx_ms_con_ssi" },
                            { name: "custrecord_efx_ms_con_qty" },
                            { name: "custrecord_efx_ms_con_w" },
                            { name: "custrecord_efx_ms_con_unit_w" },
                            { name: "custrecord_efx_ms_con_contact" },
                            { name: "custrecord_efx_ms_con_addr" },
                            { name: "custrecord_efx_ms_con_sm" },
                            { name: "custrecord_efx_ms_con_ship_cus" },
                            { name: "custrecord_efx_ms_con_act" },
                            { name: "created" },
                            {
                                name: "entityid",
                                join: "owner"
                            },
                            { name: "lastmodified" }
                        ]
                });
                var countData = 0;
                log.audit({
                    title: 'Results in searchAllContracts',
                    details: efx_ms_contract_detailSearch.runPaged().count
                });
                var myPagedResults = efx_ms_contract_detailSearch.runPaged({
                    pageSize: 1000
                });
                var thePageRanges = myPagedResults.pageRanges;
                for (var i in thePageRanges) {
                    var thepageData = myPagedResults.fetch({
                        index: thePageRanges[i].index
                    });
                    thepageData.data.forEach(function (result) {
                        if (!result.getValue({ name: 'custrecord_efx_ms_con_sm' })) {
                            countData++;
                            rolErrors.push({
                                id: result.getValue({ name: 'internalid' })
                            })
                        }
                        data.push({
                            publicacion: today,
                            producto: result.getText({ name: 'custrecord_efx_ms_con_ssi' }),
                            cantidad: result.getValue({ name: 'custrecord_efx_ms_con_qty' }),
                            peso: result.getValue({ name: 'custrecord_efx_ms_con_w' }),
                            tercero: result.getValue({ name: 'custrecord_efx_ms_con_contact' }),
                            shipping_id: result.getValue({ name: 'custrecord_efx_ms_con_sm' }),
                            usuario: result.getText({ name: 'custrecord_efx_ms_con_ship_cus' }),
                            direccion: result.getValue({ name: 'custrecord_efx_ms_con_addr' }),
                            compania: result.getText({ name: 'custrecord_efx_ms_con_sm' }),
                            activo: result.getValue({ name: 'custrecord_efx_ms_con_act' }),
                            fecha_creacion: result.getValue({ name: 'created' }),
                            creado_por: result.getValue({ name: 'entityid', join: 'owner' }),
                            actualizado: result.getValue({ name: 'lastmodified' })
                        });
                        dataRol.push({
                            custrecord_efx_ms_dre_re: '',
                            custrecord_efx_ms_dre_item: result.getValue({ name: 'custrecord_efx_ms_con_ssi' }),
                            custrecord_efx_ms_dre_sm: result.getValue({ name: 'custrecord_efx_ms_con_sm' }),
                            custrecord_efx_ms_dre_detail_contract: result.getValue({ name: 'internalid' }),
                        });
                    });
                }
                /*efx_ms_contract_detailSearch.run().each(function (result) {
                    if (!result.getValue({name: 'custrecord_efx_ms_con_sm'})) {
                        countData++;
                    }
                    data.push({
                        publicacion: today,
                        producto: result.getText({name: 'custrecord_efx_ms_con_ssi'}),
                        cantidad: result.getValue({name: 'custrecord_efx_ms_con_qty'}),
                        peso: result.getValue({name: 'custrecord_efx_ms_con_w'}),
                        tercero: result.getValue({name: 'custrecord_efx_ms_con_contact'}),
                        shipping_id: result.getValue({name: 'custrecord_efx_ms_con_sm'}),
                        usuario: result.getText({name: 'custrecord_efx_ms_con_ship_cus'}),
                        direccion: result.getValue({name: 'custrecord_efx_ms_con_addr'}),
                        compania: result.getText({name: 'custrecord_efx_ms_con_sm'}),
                        activo: result.getValue({name: 'custrecord_efx_ms_con_act'}),
                        fecha_creacion: result.getValue({name: 'created'}),
                        creado_por: result.getValue({name: 'entityid', join: 'owner'}),
                        actualizado: result.getValue({name: 'lastmodified'})
                    });
                    dataRol.push({
                        custrecord_efx_ms_dre_re: '',
                        custrecord_efx_ms_dre_item: result.getValue({name: 'custrecord_efx_ms_con_ssi'}),
                        custrecord_efx_ms_dre_sm: result.getValue({name: 'custrecord_efx_ms_con_sm'}),
                        custrecord_efx_ms_dre_detail_contract: result.getValue({name: 'internalid'}),
                    });
                    return true;
                });*/
                return { dataSL: data, dataRol: dataRol, countError: countData, rolErrors: rolErrors };
            } catch (e) {
                log.error('Error on searchAllContracts', e);
            }
        }

        function createCSV(jsonData) {
            try {
                var json = JSON.parse(jsonData);
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
                    contents: csv
                });
                fileObj.folder = -15;
                var fileId = fileObj.save();


                return { fileID: fileId };
            } catch (e) {
                log.error('Error on createCSV', e);
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

        function createRol(period, peso) {
            log.audit({title: 'Create peso', details: peso});
            try {
                var userObj = runtime.getCurrentUser();
                var recordObj = record.create({
                    type: 'customrecord_efx_ms_rol_env',
                    isDynamic: true
                });
                recordObj.setValue({
                    fieldId: 'custrecord_efx_ms_re_e',
                    value: userObj.id,
                    ignoreFieldChange: false
                });
                recordObj.setValue({
                    fieldId: 'custrecord_efx_ms_re_period',
                    value: period,
                    ignoreFieldChange: false
                });
                recordObj.setValue({
                    fieldId: 'custrecord_tkio_weight',
                    value: peso,
                    ignoreFieldChange: false
                });
                var recordId = recordObj.save({
                    enableSourcing: true,
                    ignoreMandatoryFields: false
                });
                return recordId;
            } catch (e) {
                log.error('Error on createRol', e);
            }
        }

        function createRolDetail(rolId, detailsData) {
            try {
                detailsData = JSON.parse(detailsData);
                var detailRolProcess = [];
                for (var i = 0; i < detailsData.length; i++) {
                    var recordObj = record.create({
                        type: 'customrecord_efx_ms_detail_rol_env',
                        isDynamic: true
                    });
                    recordObj.setValue({
                        fieldId: 'custrecord_efx_ms_dre_re',
                        value: rolId,
                        ignoreFieldChange: true
                    });
                    recordObj.setValue({
                        fieldId: 'custrecord_efx_ms_dre_item',
                        value: detailsData[i].item,
                        ignoreFieldChange: true
                    });
                    recordObj.setValue({
                        fieldId: 'custrecord_efx_ms_dre_sm',
                        value: detailsData[i].shipping_id,
                        ignoreFieldChange: true
                    });
                    recordObj.setValue({
                        fieldId: 'custrecord_efx_ms_dre_detail_contract',
                        value: detailsData[i].id_detail,
                        ignoreFieldChange: true
                    });
                    var recordId = recordObj.save({
                        enableSourcing: true,
                        ignoreMandatoryFields: false
                    });
                    detailRolProcess.push(recordId);
                }
                return detailRolProcess;
            } catch (e) {
                log.error('Error on createRolDetail', e);
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
                                name: "custrecord_efx_ms_con_sales_order_mirror",
                                join: "CUSTRECORD_EFX_MS_DRE_DETAIL_CONTRACT"
                            },
                            {
                                name: "formulanumeric",
                                formula: "TO_NUMBER(TO_CHAR({created},'MM'))"
                            },
                            { name: "formulatext", formula: "TO_CHAR({created},'DDMMHHMI')" }
                        ]
                });
                customrecord_efx_ms_detail_rol_envSearchObj.run().each(function (result) {
                    var itemService = result.getValue({ name: "custrecord_efx_ms_dre_item" });
                    var period = result.getValue({
                        name: "custrecord_efx_ms_con_period",
                        join: "CUSTRECORD_EFX_MS_DRE_DETAIL_CONTRACT"
                    });
                    var month = result.getValue({
                        name: "formulanumeric",
                        formula: "TO_NUMBER(TO_CHAR({created},'MM'))"
                    });
                    var detailContract = result.getValue({ name: "custrecord_efx_ms_dre_detail_contract" });
                    var prefix = result.getValue({ name: "formulatext", formula: "TO_CHAR({created},'DDMMHHMI')" });
                    var date = format.parse({
                        value: result.getValue({ name: "created" }),
                        type: format.Type.DATE
                    });
                    resultData.push({
                        item: itemService,
                        period: period,
                        month: month,
                        detailContract: detailContract,
                        prefix: prefix,
                        date_rol: new Date(date)
                    });
                    return true;
                });
                return resultData;
            } catch (e) {
                log.error('Error on getDataFromRol', e);
            }
        }

        function getAllDeailContract(detailID, periodVal, consecutive, prefix, dateRol, detailsObj) {
            try {
                log.audit({
                    title: 'getAllDeailContract Data', details: {
                        detailID: detailID,
                        period: periodVal,
                        consecutive: consecutive,
                        prefix: prefix,
                        dateRol: dateRol,
                        detailsObj: detailsObj
                    }
                });
                var details = [];
                for (var i = 0; i < detailsObj.length; i++) {
                    details.push(detailsObj[i]);
                }
                var detailData = [];
                var filters = [
                    ['internalid', search.Operator.ANYOF, details]
                ];
                log.audit({ title: 'filtros', details: filters });
                var searchDetail = search.create({
                    type: 'customrecord_efx_ms_contract_detail',
                    filters: filters,
                    columns: [
                        { name: 'internalid' },
                        { name: 'custrecord_efx_ms_con_postmark' },
                        { name: 'custrecord_efx_ms_con_sd' },
                        { name: 'custrecord_efx_ms_con_ed' },
                        { name: 'custrecord_efx_ms_con_contact' },
                        { name: 'custrecord_efx_ms_con_ship_cont' },
                        { name: 'custrecord_efx_ms_con_ship_cus' },
                        { name: 'custrecord_efx_ms_con_addr' },
                        { name: 'custrecord_efx_ms_con_des' },
                        { name: 'custrecord_efx_ms_con_qty' },
                        { name: 'custrecord_efx_ms_con_w' },
                        { name: 'custrecord_efx_ms_con_unit_w' },
                        { name: 'custrecord_efx_ms_con_period' },
                        { name: 'custrecord_efx_ms_con_sm' },
                        { name: 'custrecord_efx_ms_con_ssi' },
                        { name: 'custrecord_efx_ms_con_act' },
                        { name: 'lastmodified' },
                        { name: 'entityid', join: 'owner' },
                        { name: 'custrecord_efx_ms_con' }
                    ]
                });
                var itemDataObj = {};
                searchDetail.run().each(function (result) {
                    if (periodVal === 4 || periodVal === "4") {
                        var d = new Date(dateRol);
                        var weekC = weekCount(d.getFullYear(), d.getMonth() + 1);
                        var noWeek = null;
                        var fYear = null;
                        var monthly = false;
                        for (var j = 0; j < weekC; j++) {
                            var weeks_add = 7 * (j);
                            noWeek = new Date(d.getTime() + (24 * 60 * 60 * 1000 * weeks_add));
                            var conse = getWeek(new Date(noWeek.getFullYear(), noWeek.getMonth(), noWeek.getDate()));
                            fYear = noWeek.getFullYear();

                            var id = result.getText({
                                name: 'internalid'
                            });
                            var custrecord_efx_ms_con_postmark = result.getText({
                                name: 'custrecord_efx_ms_con_postmark'
                            });
                            var custrecord_efx_ms_con_sd = result.getValue({
                                name: 'custrecord_efx_ms_con_sd'
                            });
                            var custrecord_efx_ms_con_ed = result.getValue({
                                name: 'custrecord_efx_ms_con_ed'
                            });
                            var contact = result.getValue({ name: 'custrecord_efx_ms_con_contact' });
                            var custrecord_efx_ms_con_contact = contact;
                            var custrecord_efx_ms_con_ship_cont = result.getValue({
                                name: 'custrecord_efx_ms_con_ship_cont'
                            });
                            var custrecord_efx_ms_con_ship_cus = result.getText({
                                name: 'custrecord_efx_ms_con_ship_cus'
                            });
                            var custrecord_efx_ms_con_addr = result.getValue({
                                name: 'custrecord_efx_ms_con_addr'
                            });
                            if (custrecord_efx_ms_con_contact) {
                                var related_add = result.getValue({
                                    name: "custrecord_efx_ms_con_ship_cont"
                                });
                                custrecord_efx_ms_con_addr = getFullAddress(related_add);
                            }
                            var custrecord_efx_ms_con_des = result.getValue({ name: 'custrecord_efx_ms_con_des' });
                            var custrecord_efx_ms_con_qty = result.getValue({ name: 'custrecord_efx_ms_con_qty' });
                            var custrecord_efx_ms_con_w = result.getValue({ name: 'custrecord_efx_ms_con_w' });
                            var custrecord_efx_ms_con_unit_w = result.getText({ name: 'custrecord_efx_ms_con_unit_w' });
                            var custrecord_efx_ms_con_sm = result.getText({ name: 'custrecord_efx_ms_con_sm' });
                            var itemSSI = result.getValue({ name: 'custrecord_efx_ms_con_ssi' }).toString();
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
                            //var itemData = getInventoryItem(result.getValue({name: 'custrecord_efx_ms_con_ssi'}), period, conse);
                            var custrecord_efx_ms_con_ssi = '';
                            var formulatext_month = '';
                            var custrecord_efx_ms_con_ssi_w = '';
                            var prefixLocal = prefix;
                            /*log.audit('dates', {
                                week: new Date(noweek),
                                end_date: new Date(format.parse({value:custrecord_efx_ms_con_ed, type: format.Type.DATE}))
                            });*/
                            var custrecord_efx_ms_con_act = result.getValue({ name: 'custrecord_efx_ms_con_act' });
                            var lastmodified = result.getValue({ name: 'lastmodified' });
                            var entityid = result.getValue({ name: 'entityid', join: 'owner' });
                            var period = result.getValue({ name: 'custrecord_efx_ms_con_period' });
                            var custrecord_efx_ms_con = result.getValue({ name: 'custrecord_efx_ms_con' });
                            // if (new Date(noweek) < new Date(format.parse({value:custrecord_efx_ms_con_ed, type: format.Type.DATE}))) {
                            detailData.push({
                                id: id,
                                custrecord_efx_ms_con_postmark: custrecord_efx_ms_con_postmark,
                                custrecord_efx_ms_con_sd: custrecord_efx_ms_con_sd,
                                custrecord_efx_ms_con_ed: custrecord_efx_ms_con_ed,
                                custrecord_efx_ms_con_contact: custrecord_efx_ms_con_contact,
                                custrecord_efx_ms_con_ship_cont: custrecord_efx_ms_con_ship_cont,
                                custrecord_efx_ms_con_ship_cus: custrecord_efx_ms_con_ship_cus,
                                custrecord_efx_ms_con_addr: custrecord_efx_ms_con_addr,
                                custrecord_efx_ms_con_des: custrecord_efx_ms_con_des,
                                custrecord_efx_ms_con_qty: custrecord_efx_ms_con_qty,
                                custrecord_efx_ms_con_w: custrecord_efx_ms_con_w,
                                custrecord_efx_ms_con_unit_w: custrecord_efx_ms_con_unit_w,
                                custrecord_efx_ms_con_sm: custrecord_efx_ms_con_sm,
                                custrecord_efx_ms_con_ssi: custrecord_efx_ms_con_ssi,
                                custrecord_efx_ms_con_ssi_w: custrecord_efx_ms_con_ssi_w,
                                formulatext_month: formulatext_month,
                                prefix: prefix,
                                custrecord_efx_ms_con_act: custrecord_efx_ms_con_act,
                                lastmodified: lastmodified,
                                entityid: entityid,
                                custrecord_efx_ms_con_period: period,
                                dateRol: dateRol,
                                custrecord_efx_ms_con: custrecord_efx_ms_con,
                                consecutive: conse.toString(),
                                yearApply: fYear.toString(),
                                itemSSI: itemSSI
                            });
                            // }
                        }
                    } else {
                        var d2 = new Date(dateRol);
                        var id = result.getText({
                            name: 'internalid'
                        });
                        var custrecord_efx_ms_con_postmark = result.getText({
                            name: 'custrecord_efx_ms_con_postmark'
                        });
                        var custrecord_efx_ms_con_sd = result.getValue({
                            name: 'custrecord_efx_ms_con_sd'
                        });
                        var custrecord_efx_ms_con_ed = result.getValue({
                            name: 'custrecord_efx_ms_con_ed'
                        });
                        var contact = result.getValue({ name: 'custrecord_efx_ms_con_contact' });
                        var custrecord_efx_ms_con_contact = contact;
                        var custrecord_efx_ms_con_ship_cont = result.getValue({
                            name: 'custrecord_efx_ms_con_ship_cont'
                        });
                        var custrecord_efx_ms_con_ship_cus = result.getText({
                            name: 'custrecord_efx_ms_con_ship_cus'
                        });
                        var custrecord_efx_ms_con_addr = result.getValue({
                            name: 'custrecord_efx_ms_con_addr'
                        });
                        if (custrecord_efx_ms_con_contact) {
                            var related_add = result.getValue({
                                name: "custrecord_efx_ms_con_ship_cont"
                            });
                            log.audit('related_add', related_add);
                            custrecord_efx_ms_con_addr = getFullAddress(related_add);
                        }
                        var custrecord_efx_ms_con_des = result.getValue({ name: 'custrecord_efx_ms_con_des' });
                        var custrecord_efx_ms_con_qty = result.getValue({ name: 'custrecord_efx_ms_con_qty' });
                        var custrecord_efx_ms_con_w = result.getValue({ name: 'custrecord_efx_ms_con_w' });
                        var custrecord_efx_ms_con_unit_w = result.getText({ name: 'custrecord_efx_ms_con_unit_w' });
                        var custrecord_efx_ms_con_sm = result.getText({ name: 'custrecord_efx_ms_con_sm' });
                        // var itemData = getInventoryItem(result.getValue({name: 'custrecord_efx_ms_con_ssi'}), period, consecutive);
                        var itemSSI = result.getValue({ name: 'custrecord_efx_ms_con_ssi' }).toString();
                        monthly = true;
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
                        var formulatext_month = '';
                        var custrecord_efx_ms_con_ssi_w = '';
                        var prefixLocal = prefix;
                        var custrecord_efx_ms_con_act = result.getValue({ name: 'custrecord_efx_ms_con_act' });
                        var lastmodified = result.getValue({ name: 'lastmodified' });
                        var entityid = result.getValue({ name: 'entityid', join: 'owner' });
                        var custrecord_efx_ms_con = result.getValue({ name: 'custrecord_efx_ms_con' });
                        detailData.push({
                            id: id,
                            custrecord_efx_ms_con_postmark: custrecord_efx_ms_con_postmark,
                            custrecord_efx_ms_con_sd: custrecord_efx_ms_con_sd,
                            custrecord_efx_ms_con_ed: custrecord_efx_ms_con_ed,
                            custrecord_efx_ms_con_contact: custrecord_efx_ms_con_contact,
                            custrecord_efx_ms_con_ship_cont: custrecord_efx_ms_con_ship_cont,
                            custrecord_efx_ms_con_ship_cus: custrecord_efx_ms_con_ship_cus,
                            custrecord_efx_ms_con_addr: custrecord_efx_ms_con_addr,
                            custrecord_efx_ms_con_des: custrecord_efx_ms_con_des,
                            custrecord_efx_ms_con_qty: custrecord_efx_ms_con_qty,
                            custrecord_efx_ms_con_w: custrecord_efx_ms_con_w,
                            custrecord_efx_ms_con_unit_w: custrecord_efx_ms_con_unit_w,
                            custrecord_efx_ms_con_sm: custrecord_efx_ms_con_sm,
                            custrecord_efx_ms_con_ssi: custrecord_efx_ms_con_ssi,
                            custrecord_efx_ms_con_ssi_w: custrecord_efx_ms_con_ssi_w,
                            formulatext_month: formulatext_month,
                            prefix: prefix,
                            custrecord_efx_ms_con_act: custrecord_efx_ms_con_act,
                            lastmodified: lastmodified,
                            entityid: entityid,
                            dateRol: dateRol,
                            custrecord_efx_ms_con: custrecord_efx_ms_con,
                            consecutive: consecutive.toString(),
                            yearApply: fYear.toString(),
                            itemSSI: itemSSI
                        })
                    }
                    return true;
                });
                log.audit({ title: 'itemDataObj', details: itemDataObj });
                var itemsProcess = getItemSales(itemDataObj);
                if (itemsProcess) {
                    var kyIP = Object.keys(itemsProcess);
                    for (var ip in kyIP) {
                        for (var a in detailData) {
                            for (var b in itemsProcess[kyIP[ip]]) {
                                if (detailData[a].itemSSI === kyIP[ip]) {
                                    if (detailData[a].yearApply === itemsProcess[kyIP[ip]][b].year) {
                                        if (detailData[a].consecutive === itemsProcess[kyIP[ip]][b].consecutive) {
                                            detailData[a].custrecord_efx_ms_con_ssi = itemsProcess[kyIP[ip]][b].itemName;
                                            detailData[a].custrecord_efx_ms_con_ssi_w = itemsProcess[kyIP[ip]][b].weight;
                                            detailData[a].formulatext_month = itemsProcess[kyIP[ip]][b].period;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                log.audit({ title: 'detailData getAllDeailContract', details: detailData });
                return detailData;
            } catch (e) {
                log.error('Error on getAllDeailContract', e);
            }
        }

        function getAllDeailContract_v2(detailID, period, consecutive, prefix, dateRol) {
            try {
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
                        { name: 'custrecord_efx_ms_con_postmark' },
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
                        { name: 'custrecord_efx_ms_con' },
                        { name: 'custrecord_efx_ms_con_tp' },
                        {
                            name: "formulatext_state",
                            formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN    CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} , 'Mexico City') > 0 THEN 'CDMX'    WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} , 'México (Estado de)') > 0 THEN 'Estado de México'    ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}     END WHEN {custrecord_efx_ms_con_contact} = 'F' THEN    CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate} , 'Mexico City') > 0 THEN 'CDMX'    WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate} , 'México (Estado de)') > 0 THEN 'Estado de México'    ELSE {custrecord_efx_ms_con_ship_cus.billstate}     END END"
                        }

                    ]
                });
                var itemDataObj = {};
                log.audit({ title: 'result count contract detail', details: searchDetail.runPaged().count });
                searchDetail.run().each(function (result) {
                    if (period === 4 || period === "4") {
                        var d = new Date(dateRol);
                        var weekC = weekCount(d.getFullYear(), d.getMonth() + 1);
                        var noWeek = null;
                        var fYear = null;
                        for (var j = 0; j < weekC; j++) {
                            var weeks_add = 7 * (j);
                            noWeek = new Date(d.getTime() + (24 * 60 * 60 * 1000 * weeks_add));
                            log.audit({ title: 'no week', details: noWeek });
                            var conse = getWeek(new Date(noWeek.getFullYear(), noWeek.getMonth(), noWeek.getDate()));
                            fYear = noWeek.getFullYear();

                            var custrecord_efx_ms_con_postmark = result.getText({
                                name: 'custrecord_efx_ms_con_postmark'
                            });
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
                            var custrecord_efx_ms_con_addr = result.getValue({
                                name: 'custrecord_efx_ms_con_addr'
                            });
                            var custrecord_efx_ms_con = result.getValue({ name: 'custrecord_efx_ms_con' });
                            var custrecord_efx_ms_con_des = result.getValue({ name: 'custrecord_efx_ms_con_des' });
                            var custrecord_efx_ms_con_qty = result.getValue({ name: 'custrecord_efx_ms_con_qty' });
                            var custrecord_efx_ms_con_w = result.getValue({ name: 'custrecord_efx_ms_con_w' });
                            var custrecord_efx_ms_con_unit_w = result.getText({ name: 'custrecord_efx_ms_con_unit_w' });
                            var custrecord_efx_ms_con_sm = result.getText({ name: 'custrecord_efx_ms_con_sm' });
                            var paqueteria = custrecord_efx_ms_con_sm;
                            var iOF = custrecord_efx_ms_con_sm.toUpperCase().indexOf('SEPOMEX ');
                            if (iOF >= 0) {
                                paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                            }
                            var custrecord_efx_ms_con_tp = result.getText({ name: 'custrecord_efx_ms_con_tp' });
                            var itemSSI = result.getValue({ name: 'custrecord_efx_ms_con_ssi' }).toString();
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
                            /*log.audit('dates', {
                                week: new Date(noweek),
                                end_date: new Date(format.parse({value:custrecord_efx_ms_con_ed, type: format.Type.DATE}))
                            });*/
                            var cpCustom = result.getValue({
                                name: "formulatext",
                                formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billzipcode}  END",
                                sort: search.Sort.ASC
                            });
                            var state = result.getValue({
                                name: "formulatext_state",
                                formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN    CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} , 'Mexico City') > 0 THEN 'CDMX'    WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} , 'México (Estado de)') > 0 THEN 'Estado de México'    ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}     END WHEN {custrecord_efx_ms_con_contact} = 'F' THEN    CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate} , 'Mexico City') > 0 THEN 'CDMX'    WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate} , 'México (Estado de)') > 0 THEN 'Estado de México'    ELSE {custrecord_efx_ms_con_ship_cus.billstate}     END END"
                            });
                            // if (new Date(noweek) < new Date(format.parse({value:custrecord_efx_ms_con_ed, type: format.Type.DATE}))) {
                            detailData.push({
                                custrecord_efx_ms_con_postmark: custrecord_efx_ms_con_postmark,
                                custrecord_efx_ms_con_sd: custrecord_efx_ms_con_sd,
                                custrecord_efx_ms_con_ed: custrecord_efx_ms_con_ed,
                                custrecord_efx_ms_con_contact: custrecord_efx_ms_con_contact,
                                custrecord_efx_ms_con_ship_cont: custrecord_efx_ms_con_ship_cont,
                                custrecord_efx_ms_con_ship_cus: custrecord_efx_ms_con_ship_cus,
                                custrecord_efx_ms_con_addr: custrecord_efx_ms_con_addr,
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
                                custrecord_efx_ms_con_period: custrecord_efx_ms_con_period,
                                custrecord_efx_ms_con_tp: custrecord_efx_ms_con_tp,
                                consecutive: conse.toString(),
                                yearApply: fYear.toString(),
                                itemSSI: itemSSI,
                                iOF: iOF,
                                paqueteria: paqueteria
                            });
                            // }
                        }
                    } else {
                        var d2 = new Date(dateRol);
                        var custrecord_efx_ms_con_postmark = result.getText({
                            name: 'custrecord_efx_ms_con_postmark'
                        });
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
                        var custrecord_efx_ms_con_addr = result.getValue({
                            name: 'custrecord_efx_ms_con_addr'
                        });
                        var custrecord_efx_ms_con = result.getValue({ name: 'custrecord_efx_ms_con' });
                        var custrecord_efx_ms_con_des = result.getValue({ name: 'custrecord_efx_ms_con_des' });
                        var custrecord_efx_ms_con_qty = result.getValue({ name: 'custrecord_efx_ms_con_qty' });
                        var custrecord_efx_ms_con_w = result.getValue({ name: 'custrecord_efx_ms_con_w' });
                        var custrecord_efx_ms_con_unit_w = result.getText({ name: 'custrecord_efx_ms_con_unit_w' });
                        var custrecord_efx_ms_con_sm = result.getText({ name: 'custrecord_efx_ms_con_sm' });
                        var paqueteria = custrecord_efx_ms_con_sm;
                        var iOF = custrecord_efx_ms_con_sm.toUpperCase().indexOf('SEPOMEX ');
                        if (iOF >= 0) {
                            paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                        }
                        var custrecord_efx_ms_con_tp = result.getText({ name: 'custrecord_efx_ms_con_tp' });
                        var itemSSI = result.getValue({ name: 'custrecord_efx_ms_con_ssi' }).toString();
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
                            formula: "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN    CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} , 'Mexico City') > 0 THEN 'CDMX'    WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} , 'México (Estado de)') > 0 THEN 'Estado de México'    ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}     END WHEN {custrecord_efx_ms_con_contact} = 'F' THEN    CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate} , 'Mexico City') > 0 THEN 'CDMX'    WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate} , 'México (Estado de)') > 0 THEN 'Estado de México'    ELSE {custrecord_efx_ms_con_ship_cus.billstate}     END END"
                        });
                        var custrecord_efx_ms_con_period = period;
                        detailData.push({
                            custrecord_efx_ms_con_postmark: custrecord_efx_ms_con_postmark,
                            custrecord_efx_ms_con_sd: custrecord_efx_ms_con_sd,
                            custrecord_efx_ms_con_ed: custrecord_efx_ms_con_ed,
                            custrecord_efx_ms_con_contact: custrecord_efx_ms_con_contact,
                            custrecord_efx_ms_con_ship_cont: custrecord_efx_ms_con_ship_cont,
                            custrecord_efx_ms_con_ship_cus: custrecord_efx_ms_con_ship_cus,
                            custrecord_efx_ms_con_addr: custrecord_efx_ms_con_addr,
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
                            custrecord_efx_ms_con_period: custrecord_efx_ms_con_period,
                            custrecord_efx_ms_con_tp: custrecord_efx_ms_con_tp,
                            consecutive: consecutive.toString(),
                            yearApply: fYear.toString(),
                            itemSSI: itemSSI,
                            paqueteria: paqueteria
                        })
                    }
                    return true;
                });
                var itemsProcess = getItemSales_v2(itemDataObj);
                log.audit({ title: 'itemsProcess', details: itemsProcess });
                if (itemsProcess) {
                    var kyIP = Object.keys(itemsProcess);
                    for (var ip in kyIP) {
                        for (var a in detailData) {
                            for (var b in itemsProcess[kyIP[ip]]) {
                                if (detailData[a].itemSSI === kyIP[ip]) {
                                    if (detailData[a].yearApply === itemsProcess[kyIP[ip]][b].year) {
                                        if (detailData[a].consecutive === itemsProcess[kyIP[ip]][b].consecutive) {
                                            detailData[a].custrecord_efx_ms_con_ssi = itemsProcess[kyIP[ip]][b].itemName;
                                            detailData[a].custrecord_efx_ms_con_ssi_w = itemsProcess[kyIP[ip]][b].weight;
                                            detailData[a].formulatext_month = itemsProcess[kyIP[ip]][b].period;
                                        }
                                    }
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

        function getItemSales_v2(itemsObj) {
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
                    var filts = [
                        ["custrecord_efx_ms_period_week.custrecord_efx_ms_period_month", search.Operator.IS, month],
                        "AND",
                        ["custrecord_efx_ms_item_sus", search.Operator.ANYOF, itemSus]
                    ];
                    log.audit({ title: 'filts in get item', details: filts });
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
                                { name: "weight", join: "custrecord_efx_ms_inventory_item" }
                            ]
                    });
                    var countData = searchObj.runPaged().count;
                    log.audit({ title: 'Count results items', details: countData });
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
                            var getWeight = result.getValue({
                                name: "weight",
                                join: "custrecord_efx_ms_inventory_item"
                            }) || 0;
                            itemInv[itemsKey[i]].push({
                                consecutive: con,
                                item: itm,
                                year: yer,
                                itemName: getItem,
                                period: getPeriod,
                                weight: getWeight
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
                            var getWeight = result.getValue({
                                name: "weight",
                                join: "custrecord_efx_ms_inventory_item"
                            }) || 0;
                            itemInv[itemsKey[i]].push({
                                consecutive: con,
                                item: itm,
                                year: yer,
                                itemName: getItem,
                                period: getPeriod,
                                weight: getWeight
                            });
                        }
                        return true;
                    });
                }

                log.audit({ title: 'gobernances use', details: runtime.getCurrentScript().getRemainingUsage() });
                return itemInv;
            } catch (e) {
                log.error('Error on getItemSales', e);
            }
        }

        function getInventoryItem(itemService, period, consecutive) {
            try {
                log.audit('Data for filter Inventory Item', {
                    itemService: itemService,
                    period: period,
                    consecutive: consecutive
                });
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
                log.audit('search result inventory', searchResult);
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

        function weekCount(year, month_number) {

            // month_number is in the range 1..12

            var firstOfMonth = new Date(year, month_number - 1, 1);
            var lastOfMonth = new Date(year, month_number, 0);

            var used = firstOfMonth.getDay() + lastOfMonth.getDate();

            return Math.ceil(used / 7);
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
                log.audit('address text', addressText);
                return addressText;
            } catch (e) {

            }
        }

        function getItemSales(itemsObj) {
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
                    log.audit({ title: 'filts', details: filts });
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
                                { name: "weight", join: "custrecord_efx_ms_inventory_item" }
                            ]
                    });
                    var countData = searchObj.runPaged().count;
                    log.audit({ title: 'Count results items', details: countData });
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
                            var getWeight = result.getValue({
                                name: "weight",
                                join: "custrecord_efx_ms_inventory_item"
                            }) || 0;
                            itemInv[itemsKey[i]].push({
                                consecutive: con,
                                item: itm,
                                year: yer,
                                itemName: getItem,
                                period: getPeriod,
                                weight: getWeight
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
                            var getWeight = result.getValue({
                                name: "weight",
                                join: "custrecord_efx_ms_inventory_item"
                            }) || 0;
                            itemInv[itemsKey[i]].push({
                                consecutive: con,
                                item: itm,
                                year: yer,
                                itemName: getItem,
                                period: getPeriod,
                                weight: getWeight
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

        /**
         * Creación de la importación CSV
         * @param fileId - Id de archivo layout a importar
         * @param mappingId - Internal Id de importación csv
         * @return {string}
         */
        function createCsvImportTaskId(fileId, mappingId) {
            try {
                var scriptTask = task.create({
                    taskType: task.TaskType.CSV_IMPORT,
                });
                scriptTask.mappingId = mappingId;
                scriptTask.importFile = file.load(fileId);
                return scriptTask.submit();
            } catch (e) {
                log.error({
                    title: 'createCsvImportTaskId - error',
                    details: e.toString()
                });
            }
        }

        function matchDate(periodo, type, year) {
            try {
                var date_now = new Date();
                log.audit({ title: 'Date now in matchDate', details: date_now });
                switch (periodo) {
                    case 1:
                    case '1':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '31/0' + periodo + '/' + year
                        }
                        break;
                    case 2:
                    case '2':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '28/0' + periodo + '/' + year
                        }
                        break;
                    case 3:
                    case '3':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '31/0' + periodo + '/' + year
                        }
                        break;
                    case 4:
                    case '4':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '30/0' + periodo + '/' + year
                        }
                        break;
                    case 5:
                    case '5':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '31/0' + periodo + '/' + year
                        }
                        break;
                    case 6:
                    case '6':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '30/0' + periodo + '/' + year
                        }
                        break;
                    case 7:
                    case '7':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '30/0' + periodo + '/' + year
                        }
                        break;
                    case 8:
                    case '8':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '31/0' + periodo + '/' + year
                        }
                        break;
                    case 9:
                    case '9':
                        if (type === 'inicio') {
                            return '01/0' + periodo + '/' + year
                        } else {
                            return '30/0' + periodo + '/' + year
                        }
                        break;
                    case 10:
                    case '10':
                        if (type === 'inicio') {
                            return '01/' + periodo + '/' + year
                        } else {
                            return '31/' + periodo + '/' + year
                        }
                        break;
                    case 11:
                    case '11':
                        if (type === 'inicio') {
                            return '01/' + periodo + '/' + year
                        } else {
                            return '30/' + periodo + '/' + year
                        }
                        break;
                    case 12:
                    case '12':
                        if (type === 'inicio') {
                            return '01/' + periodo + '/' + year
                        } else {
                            return '31/' + periodo + '/' + year
                        }
                        break;
                }
            } catch (e) {
                log.error({ title: 'Error on matchDate', details: e });
            }
        }

        return {
            onRequest: onRequest
        };

    });
