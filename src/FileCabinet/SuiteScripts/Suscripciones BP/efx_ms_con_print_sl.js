/**
 * @NApiVersion 2.x
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 */
define([
    'N/record',
    'N/search',
    'N/log',
    'N/ui/serverWidget',
    'N/format',
    'N/render',
    'N/runtime',
    'N/file',
    'N/url',
    'N/redirect',
], /**
 * @param{record} record
 * @param{search} search
 * @param{render} render
 * @param{format} format
 * @param{serverWidget} serverWidget
 * @param{runtime} runtime
 * @param{log} log
 * @param{file} file
 * @param{url} url
 * @param{redirect} redirect
 */
    function (
        record,
        search,
        log,
        serverWidget,
        format,
        render,
        runtime,
        file,
        url,
        redirect
    ) {
        var params = null
        var idsDetails = []
        var folderToSave = null

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
                log.audit({
                    title: 'gobernances use:',
                    details: runtime.getCurrentScript().getRemainingUsage(),
                })

                var response = context.response
                var request = context.request
                params = request.parameters

                folderToSave = runtime
                    .getCurrentScript()
                    .getParameter({ name: 'custscript_efx_ms_folder' })
                var logoURL = runtime
                    .getCurrentScript()
                    .getParameter({ name: 'custscript_efx_ms_logo' })
                var dataProcess = {},
                    dataArray = []

                if (params.idrol) {
                    var getDataRol = getDataFromRol(params.idrol);
                    var listaSello = getListaSelloPostal();
                    log.audit({ title: 'listaSello', details: listaSello });
                    // log.audit({title: 'getDataRol', details: getDataRol});
                    log.audit({
                        title: 'gobernances use getDataFromRol',
                        details: runtime.getCurrentScript().getRemainingUsage(),
                    })
                    log.audit({
                        title: 'Results in getDataFromRol',
                        details: getDataRol.length,
                    })
                    var detailsIDArray = []
                    for (var i = 0; i < getDataRol.length; i++) {
                        detailsIDArray.push(getDataRol[i].detailContract)
                    }
                    if (detailsIDArray.length) {
                        log.audit({
                            title: 'data for getAllDeailContract_v2',
                            details: {
                                detailsIDArray: detailsIDArray,
                                period: getDataRol[0].period,
                                month: getDataRol[0].month,
                                prefix: getDataRol[0].prefix,
                                date_rol: getDataRol[0].date_rol,
                                weight: getDataRol[0].weight,
                            },
                        })
                        var detailDataAux = getAllDeailContract_v2(
                            detailsIDArray,
                            getDataRol[0].period,
                            getDataRol[0].month,
                            getDataRol[0].prefix,
                            getDataRol[0].date_rol,
                            getDataRol[0].weight,
                            listaSello
                        )
                        var detailData = detailDataAux[0]
                        log.audit({ title: 'detailDataAux', details: detailDataAux });
                        log.audit({
                            title: 'gobernances use getAllDeailContract_v2',
                            details: runtime.getCurrentScript().getRemainingUsage(),
                        })
                        for (var j = 0; j < detailData.length; j++) {
                            dataArray.push(detailData[j])
                        }
                        log.audit({ title: 'Sort', details: 'sort data' })
                        /**
                         * Orden y agrupamiento de datos
                         */
                        var newArraySort = setSortData(dataArray)
                        log.audit({ title: 'Sort', details: 'Finish sort data' })
                        var fileObj = file.create({
                            name: 'json_content_rol.txt',
                            fileType: file.Type.PLAINTEXT,
                            contents: JSON.stringify(newArraySort),
                            encoding: file.Encoding.UTF8,
                            folder: -15,
                        })
                        var idFile = fileObj.save()
                        /*response.writeFile({file: fileObj, isInline:true});
                        response.writeLine({output: JSON.stringify({success: (idFile) ? true : false})})*/

                        dataArray = []
                        for (var newArraySortKey in newArraySort) {
                            dataArray.push(newArraySort[newArraySortKey])
                        }

                        var dataPrint = []
                        var cantidad_alt = 0
                        var peso_alt = 0
                        var periodo = null
                        for (var i = 0; i < dataArray.length; i++) {
                            if (
                                dataArray[i].custrecord_efx_ms_con_period &&
                                (dataArray[i].custrecord_efx_ms_con_period ===
                                    '4' ||
                                    dataArray[i].custrecord_efx_ms_con_period === 4)
                            ) {
                                cantidad_alt +=
                                    dataArray[i].custrecord_efx_ms_con_qty * 1
                                peso_alt =
                                    dataArray[i].custrecord_efx_ms_con_w /
                                    dataArray[i].custrecord_efx_ms_con_qty
                                periodo = dataArray[i].custrecord_efx_ms_con_period
                            }
                        }

                        customerRef = detailDataAux[1]
                        log.audit({ title: 'customerRef', details: customerRef })

                        var keys = Object.keys(detailDataAux[1])
                        log.audit({ title: 'keys', details: keys })

                        var reference = []
                        for (var i = 0; i < keys.length; i++) {
                            reference.push(customerRef[keys[i]].customerRef)
                        }
                        log.audit({ title: 'ref', details: reference })

                        if (periodo && (periodo === 4 || periodo === '4')) {
                            var contractDetail = {}
                            for (var i = 0; i < dataArray.length; i++) {
                                log.audit({
                                    title:
                                        'dataArray[' +
                                        i +
                                        '].custrecord_efx_ms_con',
                                    details: {
                                        contract:
                                            dataArray[i].custrecord_efx_ms_con,
                                        id: dataArray[i].id,
                                    },
                                })
                                if (
                                    !contractDetail.hasOwnProperty(
                                        dataArray[i].custrecord_efx_ms_con
                                    )
                                ) {
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ] = {}
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id] = {
                                        custrecord_efx_ms_con_ssi: '',
                                        custrecord_efx_ms_con_w: '',
                                    }

                                    if (
                                        contractDetail[
                                            dataArray[i].custrecord_efx_ms_con
                                        ][dataArray[i].id]
                                            .custrecord_efx_ms_con_ssi === ''
                                    ) {
                                        contractDetail[
                                            dataArray[i].custrecord_efx_ms_con
                                        ][
                                            dataArray[i].id
                                        ].custrecord_efx_ms_con_ssi =
                                            dataArray[i].custrecord_efx_ms_con_ssi
                                    }
                                    if (
                                        contractDetail[
                                            dataArray[i].custrecord_efx_ms_con
                                        ][dataArray[i].id]
                                            .custrecord_efx_ms_con_w === ''
                                    ) {
                                        contractDetail[
                                            dataArray[i].custrecord_efx_ms_con
                                        ][dataArray[i].id].custrecord_efx_ms_con_w =
                                        /* dataArray[i].custrecord_efx_ms_con_qty * */ dataArray[
                                                i
                                            ].w_ssi
                                    }
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_postmark =
                                        dataArray[i].custrecord_efx_ms_con_postmark
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_sd =
                                        dataArray[i].custrecord_efx_ms_con_sd
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_ed =
                                        dataArray[i].custrecord_efx_ms_con_ed
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_contact =
                                        dataArray[i].custrecord_efx_ms_con_contact
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_ship_cont =
                                        dataArray[i].custrecord_efx_ms_con_ship_cont
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_ship_cus =
                                        dataArray[i].custrecord_efx_ms_con_ship_cus
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_addr =
                                        dataArray[i].custrecord_efx_ms_con_cust_addr
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_des =
                                        dataArray[i].custrecord_efx_ms_con_des
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_qty =
                                        dataArray[i].custrecord_efx_ms_con_qty
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_unit_w =
                                        dataArray[i].custrecord_efx_ms_con_unit_w
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_sm =
                                        dataArray[i].custrecord_efx_ms_con_sm
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].formulatext_month =
                                        dataArray[i].formulatext_month
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].prefix = dataArray[i].prefix
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].cpCustom =
                                        dataArray[i].cpCustom
                                } else {
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id] = {
                                        custrecord_efx_ms_con_ssi: '',
                                        custrecord_efx_ms_con_w: '',
                                    }

                                    if (
                                        contractDetail[
                                            dataArray[i].custrecord_efx_ms_con
                                        ][dataArray[i].id]
                                            .custrecord_efx_ms_con_ssi === ''
                                    ) {
                                        contractDetail[
                                            dataArray[i].custrecord_efx_ms_con
                                        ][
                                            dataArray[i].id
                                        ].custrecord_efx_ms_con_ssi =
                                            dataArray[i].custrecord_efx_ms_con_ssi
                                    }
                                    if (
                                        contractDetail[
                                            dataArray[i].custrecord_efx_ms_con
                                        ][dataArray[i].id]
                                            .custrecord_efx_ms_con_w === ''
                                    ) {
                                        contractDetail[
                                            dataArray[i].custrecord_efx_ms_con
                                        ][dataArray[i].id].custrecord_efx_ms_con_w =
                                            dataArray[i].w_ssi
                                    }
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_postmark =
                                        dataArray[i].custrecord_efx_ms_con_postmark
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_sd =
                                        dataArray[i].custrecord_efx_ms_con_sd
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_ed =
                                        dataArray[i].custrecord_efx_ms_con_ed
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_contact =
                                        dataArray[i].custrecord_efx_ms_con_contact
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_ship_cont =
                                        dataArray[i].custrecord_efx_ms_con_ship_cont
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_ship_cus =
                                        dataArray[i].custrecord_efx_ms_con_ship_cus
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_addr =
                                        dataArray[i].custrecord_efx_ms_con_cust_addr
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_des =
                                        dataArray[i].custrecord_efx_ms_con_des
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_qty =
                                        dataArray[i].custrecord_efx_ms_con_qty
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][
                                        dataArray[i].id
                                    ].custrecord_efx_ms_con_unit_w =
                                        dataArray[i].custrecord_efx_ms_con_unit_w
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].custrecord_efx_ms_con_sm =
                                        dataArray[i].custrecord_efx_ms_con_sm
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].formulatext_month =
                                        dataArray[i].formulatext_month
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].prefix = dataArray[i].prefix
                                    contractDetail[
                                        dataArray[i].custrecord_efx_ms_con
                                    ][dataArray[i].id].cpCustom =
                                        dataArray[i].cpCustom
                                }
                            }
                            //log.audit({title: 'Contract Detail Data', details: contractDetail});
                            //response.write({output: JSON.stringify(contractDetail)})
                            var contAux = 0
                            var contractDetailKey = Object.keys(contractDetail)
                            for (var key in contractDetailKey) {
                                var keyIDS = Object.keys(
                                    contractDetail[contractDetailKey[key]]
                                )
                                for (var id in keyIDS) {
                                    //log.audit({title:'id contract', details: contractDetail[contractDetailKey[key]][keyIDS[id]]});
                                    dataPrint.push({
                                        custrecord_efx_ms_con_postmark:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_postmark,
                                        custrecord_efx_ms_con_sd:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_sd,
                                        custrecord_efx_ms_con_ed:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_ed,
                                        custrecord_efx_ms_con_contact:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_contact,
                                        custrecord_efx_ms_con_ship_cont:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_ship_cont,
                                        custrecord_efx_ms_con_ship_cus:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_ship_cus,
                                        custrecord_efx_ms_con_addr:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_addr,
                                        custrecord_efx_ms_con_des:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_des,
                                        custrecord_efx_ms_con_qty:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_qty,
                                        custrecord_efx_ms_con_w:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_w,
                                        custrecord_efx_ms_con_unit_w:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_unit_w,
                                        custrecord_efx_ms_con_sm:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_sm,
                                        custrecord_efx_ms_con_ssi:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].custrecord_efx_ms_con_ssi,
                                        formulatext_month:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].formulatext_month,
                                        prefix: contractDetail[
                                            contractDetailKey[key]
                                        ][keyIDS[id]].prefix,
                                        cpCustom:
                                            contractDetail[contractDetailKey[key]][
                                                keyIDS[id]
                                            ].cpCustom,
                                        customRef: reference[contAux],
                                        weight: getDataRol[0].weight
                                        /**
                                         * TODO Revisar que el datos de customeRef se esté llenando bien aqui,
                                         * TODO tomar como referencia el como se llena en el else
                                         */
                                    })
                                    contAux++
                                    log.audit({
                                        title: 'reference',
                                        details: {
                                            reference: reference,
                                            contAux: contAux,
                                            id: id,
                                            key: key,
                                        },
                                    })
                                }
                            }
                        } else {
                            for (var i = 0; i < dataArray.length; i++) {
                                dataPrint.push({
                                    custrecord_efx_ms_con_postmark:
                                        dataArray[i].custrecord_efx_ms_con_postmark,
                                    custrecord_efx_ms_con_sd:
                                        dataArray[i].custrecord_efx_ms_con_sd,
                                    custrecord_efx_ms_con_ed:
                                        dataArray[i].custrecord_efx_ms_con_ed,
                                    custrecord_efx_ms_con_contact:
                                        dataArray[i].custrecord_efx_ms_con_contact,
                                    custrecord_efx_ms_con_ship_cont:
                                        dataArray[i]
                                            .custrecord_efx_ms_con_ship_cont,
                                    custrecord_efx_ms_con_ship_cus:
                                        dataArray[i].custrecord_efx_ms_con_ship_cus,
                                    custrecord_efx_ms_con_addr:
                                        dataArray[i]
                                            .custrecord_efx_ms_con_cust_addr,
                                    custrecord_efx_ms_con_des:
                                        dataArray[i].custrecord_efx_ms_con_des,
                                    custrecord_efx_ms_con_qty:
                                        dataArray[i].custrecord_efx_ms_con_qty,
                                    custrecord_efx_ms_con_w: dataArray[i].custrecord_efx_ms_con_ssi_w,
                                    custrecord_efx_ms_con_unit_w:
                                        dataArray[i].custrecord_efx_ms_con_unit_w,
                                    custrecord_efx_ms_con_sm:
                                        dataArray[i].custrecord_efx_ms_con_sm,
                                    custrecord_efx_ms_con_ssi:
                                        dataArray[i].custrecord_efx_ms_con_ssi,
                                    formulatext_month:
                                        dataArray[i].formulatext_month,
                                    prefix: dataArray[i].prefix,
                                    cpCustom: dataArray[i].cpCustom,
                                    customRef: reference[i],
                                    weight: getDataRol[0].weight
                                })
                                log.audit({
                                    title: 'reference',
                                    details: { reference: reference, i: i },
                                })
                            }
                        }
                        dataProcess['data'] = dataPrint
                        //dataProcess['ref'] = detailDataAux[1];
                        dataProcess['logo'] = logoURL
                        log.audit({ title: 'dataProcess', details: dataProcess })

                        log.audit({
                            title: 'gobernances use print',
                            details: runtime.getCurrentScript().getRemainingUsage(),
                        })
                        log.audit({
                            title: 'Data for print',
                            details: dataPrint.length,
                        })
                        log.audit({ title: 'Data for print', details: dataPrint })

                        //response.writeLine({output: JSON.stringify(dataProcess)})
                        if (dataArray.length) {
                            log.audit({
                                title: 'Prepare to render data',
                                details: ' ',
                            })

                            log.audit({
                                title: 'Prepare to render data',
                                details: 'create',
                            })
                            var renderer = render.create()
                            // Template ID
                            var scriptObj = runtime.getCurrentScript()

                            log.audit({
                                title: 'Prepare to render data',
                                details: 'set template',
                            })
                            var templateID = parseInt(
                                scriptObj.getParameter({
                                    name: 'custscript_efx_ms_template_id',
                                })
                            )
                            renderer.setTemplateById(templateID)

                            log.audit({
                                title: 'Prepare to render data',
                                details: 'set data',
                            })
                            renderer.addCustomDataSource({
                                format: render.DataSource.OBJECT,
                                alias: 'record',
                                data: dataProcess,
                            })

                            log.audit({
                                title: 'Prepare to render data',
                                details: 'render pdf',
                            })
                            var renderFile = renderer.renderAsPdf()

                            log.audit({
                                title: 'Prepare to render data',
                                details: 'savePDF',
                            })
                            // response.writeFile(renderFile, true);
                            renderFile.folder = folderToSave
                            renderFile.name =
                                'Etiquetas_rol_' + params.idrol + '.pdf'
                            var file_id = renderFile.save()
                            log.audit({ title: 'File PDF ID', details: file_id })
                            if (file_id) {
                                var file_url = getFileURL(file_id)
                                if (file_url) {
                                    redirect.redirect({
                                        url: file_url,
                                    })
                                }
                            } else {
                                var formError = serverWidget.createForm({
                                    title: 'MS - Print Label SL',
                                })
                                formError.clientScriptModulePath =
                                    './efx_ms_del_rol_cs.js'
                                var errorFld = formError.addField({
                                    id: 'custpage_fld_error',
                                    label: 'Advertencia',
                                    type: serverWidget.FieldType.LONGTEXT,
                                })
                                errorFld.defaultValue =
                                    'No se pudo procesar su información, intente nuevamente.'
                                errorFld.updateDisplayType({
                                    displayType:
                                        serverWidget.FieldDisplayType.HIDDEN,
                                })
                                log.error('Suitelet error, function onRequest', e)

                                context.response.writePage(formError)
                            }
                        }
                    }
                }
            } catch (e) {
                var formError = serverWidget.createForm({
                    title: 'MS - Print Label SL',
                })
                formError.clientScriptModulePath = './efx_ms_del_rol_cs.js'
                var errorFld = formError.addField({
                    id: 'custpage_fld_error',
                    label: 'Advertencia',
                    type: serverWidget.FieldType.LONGTEXT,
                })
                errorFld.defaultValue = e.message
                errorFld.updateDisplayType({
                    displayType: serverWidget.FieldDisplayType.HIDDEN,
                })
                log.error('Suitelet error, function onRequest', e)

                context.response.writePage(formError)
            }
        }

        function getDataFromRol(idRol) {
            try {
                var resultData = []
                var customrecord_efx_ms_detail_rol_envSearchObj = search.create({
                    type: 'customrecord_efx_ms_detail_rol_env',
                    filters: [
                        ['custrecord_efx_ms_dre_re', search.Operator.IS, idRol],
                    ],
                    columns: [
                        { name: 'internalid' },
                        { name: 'created' },
                        { name: 'custrecord_efx_ms_dre_re' },
                        { name: 'custrecord_efx_ms_dre_item' },
                        { name: 'custrecord_efx_ms_dre_sm' },
                        { name: 'custrecord_efx_ms_dre_detail_contract' },
                        {
                            name: 'custrecord_efx_ms_re_period',
                            join: 'custrecord_efx_ms_dre_re',
                        },
                        {
                            name: 'custrecord_tkio_weight',
                            join: 'custrecord_efx_ms_dre_re',
                        },
                        {
                            name: 'custrecord_efx_ms_con_period',
                            join: 'CUSTRECORD_EFX_MS_DRE_DETAIL_CONTRACT',
                        },
                        {
                            name: 'custrecord_efx_ms_con_sales_order_mirror',
                            join: 'CUSTRECORD_EFX_MS_DRE_DETAIL_CONTRACT',
                        },
                        {
                            name: 'formulanumeric',
                            formula: "TO_NUMBER(TO_CHAR({created},'MM'))",
                        },
                        {
                            name: 'formulatext',
                            formula: "TO_CHAR({created},'DDMMHHMI')",
                        },
                        { name: "custrecord_efx_ms_re_year_apply", join: "custrecord_efx_ms_dre_re" }
                    ],
                })
                log.audit({
                    title: 'Results counts detail rol',
                    details:
                        customrecord_efx_ms_detail_rol_envSearchObj.runPaged()
                            .count,
                })
                var myPagedResults =
                    customrecord_efx_ms_detail_rol_envSearchObj.runPaged({
                        pageSize: 1000,
                    })
                var thePageRanges = myPagedResults.pageRanges
                for (var i in thePageRanges) {
                    var thepageData = myPagedResults.fetch({
                        index: thePageRanges[i].index,
                    })
                    thepageData.data.forEach(function (result) {
                        var itemService = result.getValue({
                            name: 'custrecord_efx_ms_dre_item',
                        })
                        var period = result.getValue({
                            name: 'custrecord_efx_ms_con_period',
                            join: 'CUSTRECORD_EFX_MS_DRE_DETAIL_CONTRACT',
                        })
                        /*var month = result.getValue({
                             name: "formulanumeric",
                             formula: "TO_NUMBER(TO_CHAR({created},'MM'))"
                         });*/
                        var month = result.getValue({
                            name: 'custrecord_efx_ms_re_period',
                            join: 'custrecord_efx_ms_dre_re',
                        })
                        var weight = result.getValue({
                            name: 'custrecord_tkio_weight',
                            join: 'custrecord_efx_ms_dre_re',
                        })
                        var detailContract = result.getValue({
                            name: 'custrecord_efx_ms_dre_detail_contract',
                        })
                        var prefix = result.getValue({
                            name: 'formulatext',
                            formula: "TO_CHAR({created},'DDMMHHMI')",
                        })
                        var date = format.parse({
                            value: result.getValue({ name: 'created' }),
                            type: format.Type.DATE,
                        })
                        var yearApplyRol = result.getValue({ name: "custrecord_efx_ms_re_year_apply", join: "custrecord_efx_ms_dre_re" })
                        var dateForRol = '01/' + ((period < 10) ? '0' + period : period) + '/' + yearApplyRol;
                        var d = new Date(dateForRol)
                        d.setMonth(month - 1)
                        d.setDate(1)
                        resultData.push({
                            item: itemService,
                            period: period,
                            month: month,
                            detailContract: detailContract,
                            prefix: prefix,
                            date_rol: d,
                            weight: weight
                        })
                    })
                }
                /*customrecord_efx_ms_detail_rol_envSearchObj.run().each(function (result) {
    
                     return true;
                 });*/
                return resultData
            } catch (e) {
                log.error('Error on getDataFromRol', e)
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
                            // { name: "scriptid", sort: search.Sort.ASC },
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
        function getAllDeailContract_v2(
            detailID,
            period,
            consecutive,
            prefix,
            dateRol,
            weight,
            listaSello
        ) {
            try {
                var getFirstSunday = function (month, year) {
                    try {
                        var tempDate = new Date()
                        tempDate.setHours(0, 0, 0, 0)
                        tempDate.setMonth(month)
                        tempDate.setYear(year)
                        tempDate.setDate(1)

                        var day = tempDate.getDay()
                        var toNextSun = day !== 0 ? 7 - day : 0
                        tempDate.setDate(tempDate.getDate() + toNextSun)

                        return tempDate.getDate()
                    } catch (e) {
                        log.error({ title: 'Error on getFirstSunday', details: e })
                    }
                }
                period = period * 1
                var detailData = []
                var searchDetail = search.create({
                    type: 'customrecord_efx_ms_contract_detail',
                    filters: [
                        ['internalid', search.Operator.ANYOF, detailID],
                        'AND',
                        ['isinactive', search.Operator.IS, 'F'],
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
                        {
                            name: 'custrecord_efx_ms_ca_addr',
                            join: 'custrecord_efx_ms_con_ship_cont',
                        },
                        {
                            name: 'formulatext',
                            formula:
                                "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billzipcode}  END",
                            sort: search.Sort.ASC,
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
                        {
                            name: 'custitem_efx_ms_service_weight',
                            join: 'custrecord_efx_ms_con_ssi',
                        },
                        { name: 'custrecord_efx_ms_con' },
                        { name: 'custrecord_efx_ms_con_tp' },
                        {
                            name: 'formulatext_state',
                            formula:
                                "CASE WHEN{custrecord_efx_ms_con_contact} = 'T' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END WHEN{custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE{custrecord_efx_ms_con_ship_cus.billstate} END END",
                        },
                        {
                            name: 'formulatext_cust_address',
                            formula:
                                "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'F' THEN {custrecord_efx_ms_con_ship_cus}|| ' ' ||{custrecord_efx_ms_con_ship_cus.companyname} WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'T' THEN {custrecord_efx_ms_con_ship_cus}|| ' ' ||{custrecord_efx_ms_con_ship_cus.firstname}|| ' ' ||{custrecord_efx_ms_con_ship_cus.lastname} END || ' ' || '<br />' || ' ' || 'Calle: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billaddress1}|| ' ' || 'No. ext.: ' || ' ' || ' _ ' || ' ' || 'No. int.: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billaddress3}|| ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_des}|| ' ' || '<br />' || ' ' || 'Col.: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.custentity_efx_billcolony}|| ' ' || '<br />' || ' ' || 'CP: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billzipcode}|| ' ' ||{custrecord_efx_ms_con_ship_cus.billcity}|| ' ' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE {custrecord_efx_ms_con_ship_cus.billstate} END || ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_ship_cus.billcountry} WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_atention}|| ' ' || '<br />' || ' ' || 'Calle: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_street}|| ' ' || 'No. ext.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_number}|| ' ' || 'No. int.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_int_number}|| ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_des}|| ' ' || '<br />' || ' ' || 'Col.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_col}|| ' ' || '<br />' || ' ' || 'CP: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code}|| ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_city}|| ' ' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END || ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_country} END",
                        },
                    ],
                })
                var itemDataObj = {}
                var customersObj = {}
                var idrepeat = []
                log.audit({
                    title: 'Filtros en data',
                    details: searchDetail.filters,
                })
                log.audit({
                    title: 'result count contract detail',
                    details: searchDetail.runPaged().count,
                })
                var myPagedResults = searchDetail.runPaged({
                    pageSize: 1000,
                })
                var thePageRanges = myPagedResults.pageRanges
                for (var i in thePageRanges) {
                    var thepageData = myPagedResults.fetch({
                        index: thePageRanges[i].index,
                    })
                    thepageData.data.forEach(function (result, index_val) {
                        if (period === 4 || period === '4') {
                            var d = new Date(dateRol)
                            d.setMonth(consecutive - 1)
                            d.setDate(
                                getFirstSunday(consecutive - 1, d.getFullYear())
                            )
                            var weekC = weekCount(
                                d.getFullYear(),
                                parseInt(consecutive - 1) /*d.getMonth() + 1*/
                            )
                            var noWeek = null
                            var fYear = null
                            for (var j = 0; j < weekC; j++) {
                                var weeks_add = 7 * j
                                noWeek = new Date(
                                    d.getTime() + 24 * 60 * 60 * 1000 * weeks_add
                                )
                                var conse = getWeek(
                                    new Date(
                                        noWeek.getFullYear(),
                                        noWeek.getMonth(),
                                        noWeek.getDate()
                                    )
                                )
                                if (index_val === 0) {
                                    log.audit({
                                        title: 'no week',
                                        details: { 0: noWeek, 1: conse },
                                    })
                                }
                                fYear = noWeek.getFullYear()

                                var internal_id = result.getValue({
                                    name: 'internalid',
                                })

                                var custrecord_efx_ms_con_postmark = result.getValue({ name: 'formulatext_postmark', formula: 'TO_CHAR({custrecord_efx_ms_con_postmark})', })
                                var custrecord_efx_ms_con_sd = result.getValue({
                                    name: 'custrecord_efx_ms_con_sd',
                                })
                                var custrecord_efx_ms_con_ed = result.getValue({
                                    name: 'custrecord_efx_ms_con_ed',
                                })
                                var contact = result.getValue({
                                    name: 'custrecord_efx_ms_con_contact',
                                })
                                var custrecord_efx_ms_con_contact = contact
                                var custrecord_efx_ms_con_ship_cont =
                                    result.getValue({
                                        name: 'custrecord_efx_ms_ca_addr',
                                        join: 'custrecord_efx_ms_con_ship_cont',
                                    })
                                var custrecord_efx_ms_con_ship_cus = result.getText(
                                    {
                                        name: 'custrecord_efx_ms_con_ship_cus',
                                    }
                                )
                                var con_ship_cus = result.getValue({
                                    name: 'custrecord_efx_ms_con_ship_cus',
                                })
                                if (custrecord_efx_ms_con_contact === false) {
                                    if (
                                        !customersObj.hasOwnProperty(con_ship_cus)
                                    ) {
                                        customersObj[con_ship_cus] = {}
                                        customersObj[con_ship_cus] = {
                                            id: con_ship_cus,
                                            customer:
                                                custrecord_efx_ms_con_ship_cus,
                                        }
                                    }
                                } else {
                                    idrepeat.push({
                                        id: con_ship_cus,
                                        customer: custrecord_efx_ms_con_ship_cus,
                                    })
                                }
                                var custrecord_efx_ms_con_addr = result.getValue({
                                    name: 'custrecord_efx_ms_con_addr',
                                })
                                var custrecord_efx_ms_con_cust_addr =
                                    result.getValue({
                                        name: 'formulatext_cust_address',
                                        formula:
                                            "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'F' THEN {custrecord_efx_ms_con_ship_cus}|| ' ' ||{custrecord_efx_ms_con_ship_cus.companyname} WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'T' THEN {custrecord_efx_ms_con_ship_cus}|| ' ' ||{custrecord_efx_ms_con_ship_cus.firstname}|| ' ' ||{custrecord_efx_ms_con_ship_cus.lastname} END || ' ' || '<br />' || ' ' || 'Calle: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billaddress1}|| ' ' || 'No. ext.: ' || ' ' || ' _ ' || ' ' || 'No. int.: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billaddress3}|| ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_des}|| ' ' || '<br />' || ' ' || 'Col.: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.custentity_efx_billcolony}|| ' ' || '<br />' || ' ' || 'CP: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billzipcode}|| ' ' ||{custrecord_efx_ms_con_ship_cus.billcity}|| ' ' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE {custrecord_efx_ms_con_ship_cus.billstate} END || ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_ship_cus.billcountry} WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_atention}|| ' ' || '<br />' || ' ' || 'Calle: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_street}|| ' ' || 'No. ext.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_number}|| ' ' || 'No. int.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_int_number}|| ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_des}|| ' ' || '<br />' || ' ' || 'Col.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_col}|| ' ' || '<br />' || ' ' || 'CP: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code}|| ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_city}|| ' ' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END || ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_country} END",
                                    })
                                var custrecord_efx_ms_con = result.getValue({
                                    name: 'custrecord_efx_ms_con',
                                })
                                var custrecord_efx_ms_con_des = result.getValue({
                                    name: 'custrecord_efx_ms_con_des',
                                })
                                var custrecord_efx_ms_con_qty = result.getValue({
                                    name: 'custrecord_efx_ms_con_qty',
                                })
                                var custrecord_efx_ms_con_w = (weight.length !== 0) ? weight : result.getValue({ name: 'custrecord_efx_ms_con_w', })
                                //var custrecord_efx_ms_con_w = weight;
                                var custrecord_efx_ms_con_unit_w = result.getText({
                                    name: 'custrecord_efx_ms_con_unit_w',
                                })
                                var custrecord_efx_ms_con_sm = result.getText({
                                    name: 'custrecord_efx_ms_con_sm',
                                })
                                var paqueteria = custrecord_efx_ms_con_sm;
                                var methodSent = custrecord_efx_ms_con_sm;
                                var iOF = custrecord_efx_ms_con_sm.toUpperCase().indexOf('SEPOMEX ')
                                if (iOF >= 0) {
                                    paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                                    methodSent = custrecord_efx_ms_con_sm.replace(paqueteria, '');
                                    methodSent = methodSent.replace(' ', '');

                                }

                                var custrecord_efx_ms_con_tp = result.getText({
                                    name: 'custrecord_efx_ms_con_tp',
                                })
                                var itemSSI = result
                                    .getValue({ name: 'custrecord_efx_ms_con_ssi' })
                                    .toString()
                                var itemNameSSI = result.getValue({
                                    name: 'displayname',
                                    join: 'custrecord_efx_ms_con_ssi',
                                })
                                var peso = (parseFloat(custrecord_efx_ms_con_qty) * parseFloat(custrecord_efx_ms_con_w));
                                if (weight.length !== 0) {
                                    for (var i = 0; i < listaSello.length; i++) {
                                        if (peso >= parseFloat(listaSello[i].pesoMin) && peso <= parseFloat(listaSello[i].pesoMax) && (listaSello[i].articulo).indexOf(itemNameSSI) !== -1 && (custrecord_efx_ms_con_postmark === listaSello[i].selloPostal)) {
                                            custrecord_efx_ms_con_sm = listaSello[i].methodSent;
                                            paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                                            custrecord_efx_ms_con_postmark = listaSello[i].selloPostal;

                                            // methodSent = (listaSello[i].articulo).indexOf(itemNameSSI)
                                            break;
                                        }
                                    }
                                }
                                var monthly = true
                                if (period === 4 || period === '4') {
                                    monthly = false
                                }
                                if (itemDataObj.hasOwnProperty(itemSSI) === false) {
                                    itemDataObj[itemSSI] = []
                                    itemDataObj[itemSSI].push({
                                        consecutive: conse,
                                        monthly: monthly,
                                        yearApply: fYear,
                                    })
                                } else {
                                    itemDataObj[itemSSI].push({
                                        consecutive: conse,
                                        monthly: monthly,
                                        yearApply: fYear,
                                    })
                                }
                                var itemData = ''
                                var custrecord_efx_ms_con_ssi = ''
                                var custrecord_efx_ms_con_ssi_w = ''
                                var formulatext_month = ''
                                var custrecord_efx_ms_con_period = period
                                var cpCustom = result.getValue({
                                    name: 'formulatext',
                                    formula:
                                        "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billzipcode}  END",
                                    sort: search.Sort.ASC,
                                })
                                var state = result.getValue({
                                    name: 'formulatext_state',
                                    formula:
                                        "CASE WHEN{custrecord_efx_ms_con_contact} = 'T' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END WHEN{custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE{custrecord_efx_ms_con_ship_cus.billstate} END END",
                                })
                                var matchState = matchStateCode(state)
                                if (matchState != null) {
                                    state = matchState
                                }
                                var w_ssi = (weight.length !== 0) ? weight : result.getValue({ name: 'custitem_efx_ms_service_weight', join: 'custrecord_efx_ms_con_ssi', });
                                //var w_ssi = weight;

                                // if (new Date(noweek) < new Date(format.parse({value:custrecord_efx_ms_con_ed, type: format.Type.DATE}))) {
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
                                    custrecord_efx_ms_con_qty: parseInt(custrecord_efx_ms_con_qty
                                    ),
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
                                    custrecord_efx_ms_con_period:
                                        custrecord_efx_ms_con_period,
                                    custrecord_efx_ms_con_tp:
                                        custrecord_efx_ms_con_tp,
                                    consecutive: conse.toString(),
                                    yearApply: fYear.toString(),
                                    itemSSI: itemSSI,
                                    iOF: iOF,
                                    paqueteria: paqueteria,
                                    methodSent: methodSent,
                                    con_ship_cus: con_ship_cus,
                                    itemNameSSI: itemNameSSI,
                                    w_ssi: w_ssi,
                                    weight: weight
                                })
                                // }
                            }
                        } else {
                            var d2 = new Date(dateRol)

                            var internal_id = result.getValue({
                                name: 'internalid',
                            })
                            var custrecord_efx_ms_con_postmark = result.getValue({
                                name: 'formulatext_postmark',
                                formula:
                                    'TO_CHAR({custrecord_efx_ms_con_postmark})',
                            })
                            var custrecord_efx_ms_con_sd = result.getValue({
                                name: 'custrecord_efx_ms_con_sd',
                            })
                            var custrecord_efx_ms_con_ed = result.getValue({
                                name: 'custrecord_efx_ms_con_ed',
                            })
                            var contact = result.getValue({
                                name: 'custrecord_efx_ms_con_contact',
                            })
                            var custrecord_efx_ms_con_contact = contact
                            var custrecord_efx_ms_con_ship_cont = result.getValue({
                                name: 'custrecord_efx_ms_ca_addr',
                                join: 'custrecord_efx_ms_con_ship_cont',
                            })
                            var custrecord_efx_ms_con_ship_cus = result.getText({
                                name: 'custrecord_efx_ms_con_ship_cus',
                            })
                            var con_ship_cus = result.getValue({
                                name: 'custrecord_efx_ms_con_ship_cus',
                            })
                            if (custrecord_efx_ms_con_contact === false) {
                                if (!customersObj.hasOwnProperty(con_ship_cus)) {
                                    customersObj[con_ship_cus] = {}
                                    customersObj[con_ship_cus] = {
                                        id: con_ship_cus,
                                        customer: custrecord_efx_ms_con_ship_cus,
                                    }
                                }
                            } else {
                                idrepeat.push({
                                    id: con_ship_cus,
                                    customer: custrecord_efx_ms_con_ship_cus,
                                })
                            }
                            var custrecord_efx_ms_con_addr = result.getValue({
                                name: 'custrecord_efx_ms_con_addr',
                            })
                            var custrecord_efx_ms_con_cust_addr = result.getValue({
                                name: 'formulatext_cust_address',
                                formula:
                                    "CASE WHEN {custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'F' THEN {custrecord_efx_ms_con_ship_cus}|| ' ' ||{custrecord_efx_ms_con_ship_cus.companyname} WHEN {custrecord_efx_ms_con_ship_cus.isperson} = 'T' THEN {custrecord_efx_ms_con_ship_cus}|| ' ' ||{custrecord_efx_ms_con_ship_cus.firstname}|| ' ' ||{custrecord_efx_ms_con_ship_cus.lastname} END || ' ' || '<br />' || ' ' || 'Calle: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billaddress1}|| ' ' || 'No. ext.: ' || ' ' || ' _ ' || ' ' || 'No. int.: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billaddress3}|| ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_des}|| ' ' || '<br />' || ' ' || 'Col.: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.custentity_efx_billcolony}|| ' ' || '<br />' || ' ' || 'CP: ' || ' ' ||{custrecord_efx_ms_con_ship_cus.billzipcode}|| ' ' ||{custrecord_efx_ms_con_ship_cus.billcity}|| ' ' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE {custrecord_efx_ms_con_ship_cus.billstate} END || ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_ship_cus.billcountry} WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_atention}|| ' ' || '<br />' || ' ' || 'Calle: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_street}|| ' ' || 'No. ext.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_number}|| ' ' || 'No. int.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_int_number}|| ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_des}|| ' ' || '<br />' || ' ' || 'Col.: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_col}|| ' ' || '<br />' || ' ' || 'CP: ' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code}|| ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_city}|| ' ' ||CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Ciudad de México') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Cd. de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END || ' ' || '<br />' || ' ' ||{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_country} END",
                            })
                            var custrecord_efx_ms_con = result.getValue({
                                name: 'custrecord_efx_ms_con',
                            })
                            var custrecord_efx_ms_con_des = result.getValue({
                                name: 'custrecord_efx_ms_con_des',
                            })
                            var custrecord_efx_ms_con_qty = result.getValue({
                                name: 'custrecord_efx_ms_con_qty',
                            })
                            var custrecord_efx_ms_con_w = (weight.length !== 0) ? weight : result.getValue({ name: 'custrecord_efx_ms_con_w', })
                            // var custrecord_efx_ms_con_w = weight;
                            var custrecord_efx_ms_con_unit_w = result.getText({
                                name: 'custrecord_efx_ms_con_unit_w',
                            })
                            var custrecord_efx_ms_con_sm = result.getText({
                                name: 'custrecord_efx_ms_con_sm',
                            })
                            var paqueteria = custrecord_efx_ms_con_sm;
                            var methodSent = custrecord_efx_ms_con_sm;
                            var iOF = custrecord_efx_ms_con_sm
                                .toUpperCase()
                                .indexOf('SEPOMEX ')
                            if (iOF >= 0) {
                                paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '')
                                methodSent = custrecord_efx_ms_con_sm.replace(paqueteria, '');
                                methodSent = methodSent.replace(' ', '');
                            }

                            var custrecord_efx_ms_con_tp = result.getText({
                                name: 'custrecord_efx_ms_con_tp',
                            })
                            var itemSSI = result
                                .getValue({ name: 'custrecord_efx_ms_con_ssi' })
                                .toString()
                            var itemNameSSI = result.getValue({
                                name: 'displayname',
                                join: 'custrecord_efx_ms_con_ssi',
                            })
                            var peso = (parseFloat(custrecord_efx_ms_con_qty) * parseFloat(custrecord_efx_ms_con_w));
                            if (weight.length !== 0) {
                                for (var i = 0; i < listaSello.length; i++) {
                                    if (peso >= parseFloat(listaSello[i].pesoMin) && peso <= parseFloat(listaSello[i].pesoMax) && (listaSello[i].articulo).indexOf(itemNameSSI) !== -1 && (custrecord_efx_ms_con_postmark === listaSello[i].selloPostal)) {
                                        custrecord_efx_ms_con_sm = listaSello[i].methodSent;
                                        paqueteria = custrecord_efx_ms_con_sm.replace('SEPOMEX ', '');
                                        custrecord_efx_ms_con_postmark = listaSello[i].selloPostal;
                                        // methodSent = (listaSello[i].articulo).indexOf(itemNameSSI)
                                        break;
                                    }
                                }
                            }
                            var monthly = true
                            fYear = d2.getFullYear()
                            if (itemDataObj.hasOwnProperty(itemSSI) === false) {
                                itemDataObj[itemSSI] = []
                                itemDataObj[itemSSI].push({
                                    consecutive: consecutive,
                                    monthly: monthly,
                                    yearApply: fYear,
                                })
                            } else {
                                itemDataObj[itemSSI].push({
                                    consecutive: consecutive,
                                    monthly: monthly,
                                    yearApply: fYear,
                                })
                            }
                            var custrecord_efx_ms_con_ssi = ''
                            var custrecord_efx_ms_con_ssi_w = ''
                            var formulatext_month = ''
                            var cpCustom = result.getValue({
                                name: 'formulatext',
                                formula:
                                    "CASE WHEN {custrecord_efx_ms_con_contact} = 'T' THEN {custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_zip_code} WHEN {custrecord_efx_ms_con_contact} = 'F' THEN {custrecord_efx_ms_con_ship_cus.billzipcode}  END",
                                sort: search.Sort.ASC,
                            })
                            var state = result.getValue({
                                name: 'formulatext_state',
                                formula:
                                    "CASE WHEN{custrecord_efx_ms_con_contact} = 'T' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'Mexico City') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state}, 'MEX') > 0 THEN 'Estado de México' ELSE{custrecord_efx_ms_con_ship_cont.custrecord_efx_ms_ca_state} END WHEN{custrecord_efx_ms_con_contact} = 'F' THEN CASE WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CDMX') > 0 THEN 'Ciudad de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'México (Estado de)') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MEX') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'Estado de México') > 0 THEN 'Estado de México' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'AGS') > 0 THEN 'Aguascalientes' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BC') > 0 THEN 'Baja California' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'BCS') > 0 THEN 'Baja California Sur' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CAM') > 0 THEN 'Campeche' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIS') > 0 THEN 'Chiapas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'CHIH') > 0 THEN 'Chihuahua' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COAH') > 0 THEN 'Coahuila' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'COL') > 0 THEN 'Colima' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'DGO') > 0 THEN 'Durango' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GTO') > 0 THEN 'Guanajuato' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'GRO') > 0 THEN 'Guerrero' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'HGO') > 0 THEN 'Hidalgo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'JAL') > 0 THEN 'Jalisco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MICH') > 0 THEN 'Michoacán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'MOR') > 0 THEN 'Morelos' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NAY') > 0 THEN 'Nayarit' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'NL') > 0 THEN 'Nuevo León' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'OAX') > 0 THEN 'Oaxaca' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'PUE') > 0 THEN 'Puebla' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QRO') > 0 THEN 'Querétaro' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'QROO') > 0 THEN 'Quintana Roo' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SLP') > 0 THEN 'San Luis Potosí' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SIN') > 0 THEN 'Sinaloa' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'SON') > 0 THEN 'Sonora' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAB') > 0 THEN 'Tabasco' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TAMPS') > 0 THEN 'Tamaulipas' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'TLAX') > 0 THEN 'Tlaxcala' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'VER') > 0 THEN 'Veracruz' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'YUC') > 0 THEN 'Yucatán' WHEN INSTR({custrecord_efx_ms_con_ship_cus.billstate}, 'ZAC') > 0 THEN 'Zacatecas' ELSE{custrecord_efx_ms_con_ship_cus.billstate} END END",
                            })
                            var matchState = matchStateCode(state)
                            if (matchState != null) {
                                state = matchState
                            }
                            var custrecord_efx_ms_con_period = period
                            var w_ssi = (weight.length !== 0) ? weight : result.getValue({ name: 'custitem_efx_ms_service_weight', join: 'custrecord_efx_ms_con_ssi', })
                            // var w_ssi = weight;
                            detailData.push({
                                id: internal_id,
                                custrecord_efx_ms_con_postmark:
                                    custrecord_efx_ms_con_postmark,
                                custrecord_efx_ms_con_sd: custrecord_efx_ms_con_sd,
                                custrecord_efx_ms_con_ed: custrecord_efx_ms_con_ed,
                                custrecord_efx_ms_con_contact:
                                    custrecord_efx_ms_con_contact,
                                custrecord_efx_ms_con_ship_cont:
                                    custrecord_efx_ms_con_ship_cont,
                                custrecord_efx_ms_con_ship_cus:
                                    custrecord_efx_ms_con_ship_cus,
                                custrecord_efx_ms_con_addr:
                                    custrecord_efx_ms_con_addr,
                                custrecord_efx_ms_con_cust_addr:
                                    custrecord_efx_ms_con_cust_addr,
                                custrecord_efx_ms_con_des:
                                    custrecord_efx_ms_con_des,
                                custrecord_efx_ms_con_qty: parseInt(
                                    custrecord_efx_ms_con_qty
                                ),
                                custrecord_efx_ms_con_w: custrecord_efx_ms_con_w,
                                custrecord_efx_ms_con_unit_w:
                                    custrecord_efx_ms_con_unit_w,
                                custrecord_efx_ms_con_sm: custrecord_efx_ms_con_sm,
                                custrecord_efx_ms_con_ssi:
                                    custrecord_efx_ms_con_ssi,
                                custrecord_efx_ms_con_ssi_w:
                                    custrecord_efx_ms_con_ssi_w,
                                formulatext_month: formulatext_month,
                                prefix: prefix,
                                cpCustom: cpCustom,
                                state: state,
                                custrecord_efx_ms_con: custrecord_efx_ms_con,
                                custrecord_efx_ms_con_period:
                                    custrecord_efx_ms_con_period,
                                custrecord_efx_ms_con_tp: custrecord_efx_ms_con_tp,
                                consecutive: consecutive.toString(),
                                yearApply: fYear.toString(),
                                itemSSI: itemSSI,
                                paqueteria: paqueteria,
                                methodSent: methodSent,
                                con_ship_cus: con_ship_cus,
                                itemNameSSI: itemNameSSI,
                                w_ssi: w_ssi,
                                weight: weight
                            })
                        }
                    })
                }
                var itemsProcess = getItemSales(itemDataObj, weight)
                log.audit({ title: 'itemsProcess', details: itemsProcess })
                if (itemsProcess) {
                    var kyIP = Object.keys(itemsProcess)
                    for (var ip in kyIP) {
                        for (var a in detailData) {
                            for (var b in itemsProcess[kyIP[ip]]) {
                                if (detailData[a].itemSSI === kyIP[ip]) {
                                    if (
                                        detailData[a].yearApply ===
                                        itemsProcess[kyIP[ip]][b].year
                                    ) {
                                        if (
                                            detailData[a].consecutive ===
                                            itemsProcess[kyIP[ip]][b].consecutive
                                        ) {
                                            detailData[a].custrecord_efx_ms_con_ssi = itemsProcess[kyIP[ip]][b].nameInventory
                                            detailData[a].custrecord_efx_ms_con_ssi_w = detailData[a].custrecord_efx_ms_con_w//itemsProcess[kyIP[ip]][b].weight
                                            detailData[a].formulatext_month = itemsProcess[kyIP[ip]][b].period
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                log.audit({ title: 'id repeat', details: idrepeat })
                log.audit({ title: 'id repeat', details: idrepeat.length })
                var customerProcess = getCustomerExtNumber(customersObj)
                if (customerProcess) {
                    var keyCustomer = Object.keys(customerProcess)
                    for (var cust in keyCustomer) {
                        for (var a in detailData) {
                            if (
                                detailData[a].custrecord_efx_ms_con_contact ===
                                false
                            ) {
                                if (
                                    detailData[a].con_ship_cus === keyCustomer[cust]
                                ) {
                                    detailData[a].custrecord_efx_ms_con_cust_addr =
                                        detailData[
                                            a
                                        ].custrecord_efx_ms_con_cust_addr.replace(
                                            ' _ ',
                                            customerProcess[keyCustomer[cust]]
                                                .streetnum + ' '
                                        )
                                }
                            }
                        }
                    }
                }
                var arregloGLobal = [detailData, customerProcess]

                return arregloGLobal
                // return detailData;
            } catch (e) {
                log.error('Error on getAllDeailContract', e)
            }
        }

        function getItemSales(itemsObj, weight) {
            try {
                var itemInv = {}
                var month,
                    itemSus,
                    consecutives = []
                var itemsKey = Object.keys(itemsObj)
                for (var i in itemsKey) {
                    itemSus = itemsKey[i]
                    month = itemsObj[itemsKey[i]][0].monthly ? 'T' : 'F'
                    for (var j in itemsObj[itemsKey[i]]) {
                        if (
                            consecutives.indexOf(
                                itemsObj[itemsKey[i]][j].consecutive
                            ) < 0
                        ) {
                            consecutives.push(
                                parseInt(itemsObj[itemsKey[i]][j].consecutive)
                            )
                        }
                    }
                    var minimo = Math.min.apply(null, consecutives)
                    var limite = Math.max.apply(null, consecutives)
                    var filts = [
                        [
                            'custrecord_efx_ms_period_week.custrecord_efx_ms_period_month',
                            search.Operator.IS,
                            month,
                        ],
                        'AND',
                        [
                            'custrecord_efx_ms_item_sus',
                            search.Operator.ANYOF,
                            itemSus,
                        ],
                    ]
                    log.audit({ title: 'filts in get item', details: filts })
                    var searchObj = search.create({
                        type: 'customrecord_efx_ms_sus_related',
                        filters: filts,
                        columns: [
                            { name: 'internalid' },
                            { name: 'custrecord_efx_ms_item_sus' },
                            { name: 'custrecord_efx_ms_inventory_item' },
                            {
                                name: 'custrecord_efx_ms_consecutive',
                                join: 'custrecord_efx_ms_period_week',
                                sort: search.Sort.ASC,
                            },
                            {
                                name: 'custrecord_efx_ms_year_apply',
                                sort: search.Sort.ASC,
                            },
                            { name: 'custrecord_efx_ms_period_week' },
                            {
                                name: 'weight',
                                join: 'custrecord_efx_ms_inventory_item',
                            },
                            {
                                name: 'displayname',
                                join: 'custrecord_efx_ms_inventory_item',
                            },
                        ],
                    })
                    var countData = searchObj.runPaged().count
                    log.audit({ title: 'Count results items', details: countData })
                    searchObj.run().each(function (result) {
                        if (itemInv.hasOwnProperty(itemsKey[i])) {
                            var con = result.getValue({
                                name: 'custrecord_efx_ms_consecutive',
                                join: 'custrecord_efx_ms_period_week',
                                sort: search.Sort.ASC,
                            })
                            var itm = result.getValue({
                                name: 'custrecord_efx_ms_inventory_item',
                            })
                            var yer = result.getValue({
                                name: 'custrecord_efx_ms_year_apply',
                                sort: search.Sort.ASC,
                            })
                            var getItem = result.getText({
                                name: 'custrecord_efx_ms_inventory_item',
                            })
                            var getPeriod = result.getText({
                                name: 'custrecord_efx_ms_period_week',
                            })
                            //var getWeight = weight || 0;
                            var getWeight = (weight.length !== 0) ? weight : result.getValue({ name: 'weight', join: 'custrecord_efx_ms_inventory_item', }) || 0;
                            var nameInventory = result.getValue({
                                name: 'displayname',
                                join: 'custrecord_efx_ms_inventory_item',
                            })
                            itemInv[itemsKey[i]].push({
                                consecutive: con,
                                item: itm,
                                year: yer,
                                itemName: getItem,
                                period: getPeriod,
                                weight: getWeight,
                                nameInventory: nameInventory,
                            })
                        } else {
                            itemInv[itemsKey[i]] = []
                            var con = result.getValue({
                                name: 'custrecord_efx_ms_consecutive',
                                join: 'custrecord_efx_ms_period_week',
                                sort: search.Sort.ASC,
                            })
                            var itm = result.getValue({
                                name: 'custrecord_efx_ms_inventory_item',
                            })
                            var yer = result.getValue({
                                name: 'custrecord_efx_ms_year_apply',
                                sort: search.Sort.ASC,
                            })
                            var getItem = result.getText({
                                name: 'custrecord_efx_ms_inventory_item',
                            })
                            var getPeriod = result.getText({
                                name: 'custrecord_efx_ms_period_week',
                            })
                            var getWeight = (weight.length !== 0) ? weight : result.getValue({ name: 'weight', join: 'custrecord_efx_ms_inventory_item', }) || 0
                            var nameInventory = result.getValue({
                                name: 'displayname',
                                join: 'custrecord_efx_ms_inventory_item',
                            })
                            itemInv[itemsKey[i]].push({
                                consecutive: con,
                                item: itm,
                                year: yer,
                                itemName: getItem,
                                period: getPeriod,
                                weight: getWeight,
                                nameInventory: nameInventory,
                            })
                        }
                        return true
                    })
                }

                log.audit({
                    title: 'gobernances use',
                    details: runtime.getCurrentScript().getRemainingUsage(),
                })
                return itemInv
            } catch (e) {
                log.error('Error on getItemSales', e)
            }
        }

        function weekCount(year, month_number) {
            // month_number is in the range 1..12

            var firstOfMonth = new Date(year, month_number - 1, 1)
            var lastOfMonth = new Date(year, month_number, 0)

            var used = firstOfMonth.getDay() + lastOfMonth.getDate()

            return Math.ceil(used / 7)
        }

        function getWeek(d) {
            var oneJan = new Date(d.getFullYear(), 0, 1)

            // calculating number of days
            //in given year before given date

            var numberOfDays = Math.floor((d - oneJan) / (24 * 60 * 60 * 1000))

            // adding 1 since this.getDay()
            //returns value starting from 0

            return Math.ceil((d.getDay() + 1 + numberOfDays) / 7)
        }

        function getCustomerExtNumber(customersID) {
            try {
                var customerData = {}
                var customerKey = Object.keys(customersID),
                    ids = []
                var customersRepeat = []
                for (var i in customerKey) {
                    ids.push(customersID[customerKey[i]].id)
                }
                log.audit({ title: 'customers repeat', details: customersRepeat })
                log.audit({ title: 'customersID', details: customersID })
                log.audit({
                    title: 'customersID lenght',
                    details: customerKey.length,
                })
                log.audit({ title: 'id', details: ids })
                var filt = [
                    [
                        'internalid',
                        search.Operator.ANYOF,
                        ids.length ? ids : '@NONE@',
                    ],
                    'AND',
                    ['billingaddress.isdefaultbilling', search.Operator.IS, 'T'],
                    /* "AND",
                     ["address.custrecordbpr_referenciadireccion", search.Operator.ISNOTEMPTY, ""] */
                    /* "AND",
                     ["address.custrecordbpr_referenciadireccion", search.Operator.ISNOTEMPTY, ""] */
                ]
                log.audit({ title: 'Filters customer', details: filt })
                var customerSearchObj = search.create({
                    type: search.Type.CUSTOMER,
                    filters: filt,
                    columns: [
                        { name: 'custrecord_streetnum', join: 'billingAddress' },
                        { name: 'entityid' },
                        { name: 'internalid' },
                        // {name: "custrecordbpr_referenciadireccion", join: "Address"}
                        {
                            name: 'custrecordbpr_referenciadireccion',
                            join: 'billingAddress',
                            label: 'Referencia',
                        },
                    ],
                })
                var searchResultCount = customerSearchObj.runPaged().count
                log.audit('customerSearchObj result count', searchResultCount)
                if (searchResultCount > 0) {
                    customerSearchObj.run().each(function (result) {
                        var streetnum = result.getValue({
                            name: 'custrecord_streetnum',
                            join: 'billingAddress',
                        })
                        var customer = result.getValue({ name: 'entityid' })
                        var cust_id = result.getValue({ name: 'internalid' })
                        var cust_address = result.getValue({
                            name: 'custrecordbpr_referenciadireccion',
                            join: 'billingAddress',
                            label: 'Referencia',
                        })

                        if (!customerData.hasOwnProperty(cust_id)) {
                            customerData[cust_id] = {}
                            customerData[cust_id] = {
                                id: cust_id,
                                customer: customer,
                                streetnum: streetnum,
                                customerRef: cust_address,
                            }
                        } else {
                            customerData[cust_id] = {
                                id: cust_id,
                                customer: customer,
                                streetnum: streetnum,
                                customerRef: cust_address,
                            }
                        }
                        return true
                    })
                    log.audit({
                        title: 'CustomerData',
                        details: JSON.stringify(customerData),
                    })
                }

                return customerData
            } catch (e) {
                log.error({ title: 'Error on getCustomerExtNumber', details: e })
            }
        }

        function matchStateCode(state) {
            var states = {
                AGS: 'Aguascalientes',
                BC: 'Baja California',
                BCS: 'Baja California Sur',
                CAM: 'Campeche',
                CHIS: 'Chiapas',
                CHIH: 'Chihuahua',
                COAH: 'Coahuila',
                COL: 'Colima',
                DGO: 'Durango',
                GTO: 'Guanajuato',
                GRO: 'Guerrero',
                HGO: 'Hidalgo',
                JAL: 'Jalisco',
                CDMX: 'CDMX',
                MICH: 'Michoacán',
                MOR: 'Morelos',
                MEX: 'Estado de México',
                NAY: 'Nayarit',
                NL: 'Nuevo León',
                OAX: 'Oaxaca',
                PUE: 'Puebla',
                QRO: 'Querétaro',
                QROO: 'Quintana Roo',
                SLP: 'San Luis Potosí',
                SIN: 'Sinaloa',
                SON: 'Sonora',
                TAB: 'Tabasco',
                TAMPS: 'Tamaulipas',
                TLAX: 'Tlaxcala',
                VER: 'Veracruz',
                YUC: 'Yucatán',
                ZAC: 'Zacatecas',
            }
            try {
                if (state != '') {
                    if (states.hasOwnProperty(state)) {
                        return states[state].toUpperCase()
                    } else {
                        return null
                    }
                } else {
                    return null
                }
            } catch (e) {
                log.error({ title: 'Error on matchStateCode', details: e })
                return null
            }
        }

        function setSortData(dataArray) {
            try {
                dataArray.sort(function (a, b) {
                    if (a.custrecord_efx_ms_con_tp > b.custrecord_efx_ms_con_tp) {
                        return 1
                    }
                    if (a.custrecord_efx_ms_con_tp < b.custrecord_efx_ms_con_tp) {
                        return -1
                    }
                    return 0
                })
                var newArraySort = {}
                for (var i in dataArray) {
                    if (
                        newArraySort.hasOwnProperty(
                            dataArray[i].custrecord_efx_ms_con_tp
                        ) === false
                    ) {
                        newArraySort[dataArray[i].custrecord_efx_ms_con_tp] = []
                        newArraySort[dataArray[i].custrecord_efx_ms_con_tp].push(
                            dataArray[i]
                        )
                    } else {
                        newArraySort[dataArray[i].custrecord_efx_ms_con_tp].push(
                            dataArray[i]
                        )
                    }
                }
                var extranjero = {}
                if (
                    newArraySort.hasOwnProperty('Suscripción extranjero') === true
                ) {
                    var dataExtran = newArraySort['Suscripción extranjero']
                    dataExtran.sort(function (a, b) {
                        if (
                            a.custrecord_efx_ms_con_sm > b.custrecord_efx_ms_con_sm
                        ) {
                            return 1
                        }
                        if (
                            a.custrecord_efx_ms_con_sm < b.custrecord_efx_ms_con_sm
                        ) {
                            return -1
                        }
                        return 0
                    })
                    for (var i in dataExtran) {
                        if (
                            dataExtran[i].custrecord_efx_ms_con_sm.toUpperCase() ===
                            'SEPOMEX  AEREO' ||
                            dataExtran[i].custrecord_efx_ms_con_sm.toUpperCase() ===
                            'SEPOMEX AEREO CERTIFCIADO'
                        ) {
                            if (
                                extranjero.hasOwnProperty(
                                    dataExtran[i].custrecord_efx_ms_con_sm
                                )
                            ) {
                                extranjero[
                                    dataExtran[i].custrecord_efx_ms_con_sm
                                ].push(dataExtran[i])
                            } else {
                                extranjero[dataExtran[i].custrecord_efx_ms_con_sm] =
                                    []
                                extranjero[
                                    dataExtran[i].custrecord_efx_ms_con_sm
                                ].push(dataExtran[i])
                            }
                        }
                    }
                    var keysExtranjero = Object.keys(extranjero)
                    var objTemp = []
                    for (var i in keysExtranjero) {
                        extranjero[keysExtranjero[i]].sort(function (a, b) {
                            if (
                                a.custrecord_efx_ms_con_qty >
                                b.custrecord_efx_ms_con_qty
                            ) {
                                return 1
                            }
                            if (
                                a.custrecord_efx_ms_con_qty <
                                b.custrecord_efx_ms_con_qty
                            ) {
                                return -1
                            }
                            return 0
                        })
                    }
                    for (var key in keysExtranjero) {
                        for (var i in extranjero[keysExtranjero[key]]) {
                            objTemp.push(extranjero[keysExtranjero[key]][i])
                        }
                    }
                    newArraySort['Suscripción extranjero'] = objTemp
                }
                var nacional = {}
                if (newArraySort.hasOwnProperty('Suscripción nacional') === true) {
                    var nacData = newArraySort['Suscripción nacional']
                    nacData.sort(function (a, b) {
                        if (
                            a.custrecord_efx_ms_con_sm > b.custrecord_efx_ms_con_sm
                        ) {
                            return 1
                        }
                        if (
                            a.custrecord_efx_ms_con_sm < b.custrecord_efx_ms_con_sm
                        ) {
                            return -1
                        }
                        return 0
                    })
                    for (var i in nacData) {
                        if (
                            nacional.hasOwnProperty(
                                nacData[i].custrecord_efx_ms_con_sm
                            )
                        ) {
                            nacional[nacData[i].custrecord_efx_ms_con_sm].push(
                                nacData[i]
                            )
                        } else {
                            nacional[nacData[i].custrecord_efx_ms_con_sm] = []
                            nacional[nacData[i].custrecord_efx_ms_con_sm].push(
                                nacData[i]
                            )
                        }
                    }
                    var keyNac = Object.keys(nacional)
                    var nacionalFilter = {}
                    for (var key in keyNac) {
                        if (keyNac[key] === 'SEPOMEX ORDINARIO') {
                            nacionalFilter[keyNac[key]] = nacional[keyNac[key]]
                        }
                        if (keyNac[key] === 'SEPOMEX CERTIFICADO') {
                            nacionalFilter[keyNac[key]] = nacional[keyNac[key]]
                        }

                        if (keyNac[key].indexOf('LIBRERÍA') === 0) {
                            if (
                                nacionalFilter.hasOwnProperty('Librerías') === true
                            ) {
                                nacionalFilter['Librerías'][keyNac[key]] =
                                    nacional[keyNac[key]]
                            } else {
                                nacionalFilter['Librerías'] = {}
                                nacionalFilter['Librerías'][keyNac[key]] =
                                    nacional[keyNac[key]]
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

                        if (
                            keyNac[key] != 'SEPOMEX ORDINARIO' &&
                            keyNac[key] != 'SEPOMEX CERTIFICADO' &&
                            keyNac[key].indexOf('LIBRERÍA') < 0 &&
                            keyNac[key] != 'SEPOMEX PAQ-POST' &&
                            keyNac[key] != 'DHL NORMAL NACIONAL' &&
                            keyNac[key] != 'FEDEX NACIONAL' &&
                            keyNac[key] != 'REDPACK NACIONAL'
                        ) {
                            if (
                                nacionalFilter.hasOwnProperty('Sin Logica') === true
                            ) {
                                nacionalFilter['Sin Logica'][keyNac[key]] =
                                    nacional[keyNac[key]]
                            } else {
                                nacionalFilter['Sin Logica'] = {}
                                nacionalFilter['Sin Logica'][keyNac[key]] =
                                    nacional[keyNac[key]]
                            }
                        }
                    }
                    if (nacionalFilter.hasOwnProperty('Librerías')) {
                        var librerias = nacionalFilter['Librerías']
                        var keyLib = Object.keys(librerias)
                        for (var i in keyLib) {
                            librerias[keyLib[i]].sort(function (a, b) {
                                if (
                                    a.custrecord_efx_ms_con_qty >
                                    b.custrecord_efx_ms_con_qty
                                ) {
                                    return 1
                                }
                                if (
                                    a.custrecord_efx_ms_con_qty <
                                    b.custrecord_efx_ms_con_qty
                                ) {
                                    return -1
                                }
                                return 0
                            })
                        }
                        var temp = []
                        for (var i in keyLib) {
                            for (var libreriaKey in librerias[keyLib[i]]) {
                                temp.push(librerias[keyLib[i]][libreriaKey])
                            }
                        }
                        nacionalFilter['Librerías'] = temp
                    }

                    if (nacionalFilter.hasOwnProperty('SEPOMEX ORDINARIO')) {
                        var ordinarioData = nacionalFilter['SEPOMEX ORDINARIO']
                        ordinarioData = orderData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty'
                        )
                        ordinarioData = groupData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty',
                            false
                        )
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(
                                ordinarioData[keysordinarioqty[k]],
                                'state'
                            )
                            ordinarioData[keysordinarioqty[k]] = groupData(
                                ordinarioData[keysordinarioqty[k]],
                                'state',
                                true
                            )
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] =
                                    orderData(
                                        ordinarioData[keysordinarioqty[k]][
                                        keysState[j]
                                        ],
                                        'cpCustom'
                                    )
                            }
                        }
                        var temp = []
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                var stateLocal =
                                    ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['SEPOMEX ORDINARIO'] = temp
                    }

                    if (nacionalFilter.hasOwnProperty('SEPOMEX CERTIFICADO')) {
                        var ordinarioData = nacionalFilter['SEPOMEX CERTIFICADO']
                        ordinarioData = orderData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty'
                        )
                        ordinarioData = groupData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty',
                            false
                        )
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(
                                ordinarioData[keysordinarioqty[k]],
                                'state'
                            )
                            ordinarioData[keysordinarioqty[k]] = groupData(
                                ordinarioData[keysordinarioqty[k]],
                                'state',
                                true
                            )
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] =
                                    orderData(
                                        ordinarioData[keysordinarioqty[k]][
                                        keysState[j]
                                        ],
                                        'cpCustom'
                                    )
                            }
                        }
                        var temp = []
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                var stateLocal =
                                    ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['SEPOMEX CERTIFICADO'] = temp
                    }

                    if (nacionalFilter.hasOwnProperty('SEPOMEX PAQ-POST')) {
                        var ordinarioData = nacionalFilter['SEPOMEX PAQ-POST']
                        ordinarioData = orderData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty'
                        )
                        ordinarioData = groupData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty',
                            false
                        )
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(
                                ordinarioData[keysordinarioqty[k]],
                                'state'
                            )
                            ordinarioData[keysordinarioqty[k]] = groupData(
                                ordinarioData[keysordinarioqty[k]],
                                'state',
                                true
                            )
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] =
                                    orderData(
                                        ordinarioData[keysordinarioqty[k]][
                                        keysState[j]
                                        ],
                                        'cpCustom'
                                    )
                            }
                        }
                        var temp = []
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                var stateLocal =
                                    ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['SEPOMEX PAQ-POST'] = temp
                    }

                    if (nacionalFilter.hasOwnProperty('DHL NORMAL NACIONAL')) {
                        var ordinarioData = nacionalFilter['DHL NORMAL NACIONAL']
                        ordinarioData = orderData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty'
                        )
                        ordinarioData = groupData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty',
                            false
                        )
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(
                                ordinarioData[keysordinarioqty[k]],
                                'state'
                            )
                            ordinarioData[keysordinarioqty[k]] = groupData(
                                ordinarioData[keysordinarioqty[k]],
                                'state',
                                true
                            )
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] =
                                    orderData(
                                        ordinarioData[keysordinarioqty[k]][
                                        keysState[j]
                                        ],
                                        'cpCustom'
                                    )
                            }
                        }
                        var temp = []
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                var stateLocal =
                                    ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['DHL NORMAL NACIONAL'] = temp
                    }

                    if (nacionalFilter.hasOwnProperty('FEDEX NACIONAL')) {
                        var ordinarioData = nacionalFilter['FEDEX NACIONAL']
                        ordinarioData = orderData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty'
                        )
                        ordinarioData = groupData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty',
                            false
                        )
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(
                                ordinarioData[keysordinarioqty[k]],
                                'state'
                            )
                            ordinarioData[keysordinarioqty[k]] = groupData(
                                ordinarioData[keysordinarioqty[k]],
                                'state',
                                true
                            )
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] =
                                    orderData(
                                        ordinarioData[keysordinarioqty[k]][
                                        keysState[j]
                                        ],
                                        'cpCustom'
                                    )
                            }
                        }
                        var temp = []
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                var stateLocal =
                                    ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['FEDEX NACIONAL'] = temp
                    }

                    if (nacionalFilter.hasOwnProperty('REDPACK NACIONAL')) {
                        var ordinarioData = nacionalFilter['REDPACK NACIONAL']
                        ordinarioData = orderData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty'
                        )
                        ordinarioData = groupData(
                            ordinarioData,
                            'custrecord_efx_ms_con_qty',
                            false
                        )
                        var keysordinarioqty = Object.keys(ordinarioData)
                        for (var k in keysordinarioqty) {
                            ordinarioData[keysordinarioqty[k]] = orderData(
                                ordinarioData[keysordinarioqty[k]],
                                'state'
                            )
                            ordinarioData[keysordinarioqty[k]] = groupData(
                                ordinarioData[keysordinarioqty[k]],
                                'state',
                                true
                            )
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                ordinarioData[keysordinarioqty[k]][keysState[j]] =
                                    orderData(
                                        ordinarioData[keysordinarioqty[k]][
                                        keysState[j]
                                        ],
                                        'cpCustom'
                                    )
                            }
                        }
                        var temp = []
                        for (var k in keysordinarioqty) {
                            var keysState = Object.keys(
                                ordinarioData[keysordinarioqty[k]]
                            )
                            for (var j in keysState) {
                                var stateLocal =
                                    ordinarioData[keysordinarioqty[k]][keysState[j]]
                                for (var m in stateLocal) {
                                    temp.push(stateLocal[m])
                                }
                            }
                        }
                        nacionalFilter['REDPACK NACIONAL'] = temp
                    }

                    newArraySort['Suscripción nacional'] = nacionalFilter
                }
                // "posisiones": ["Extranjero", "SEPOMEX ORDINARIO", "SEPOMEX CERTIFICADO", "LIBRERIAS", "SEPOMEX PAQ-POST", "DHL NORMAL NACIONAL", "FEDEX NACIONAL", "REDPACK NACIONAL"]
                var finalObj = []
                if (newArraySort.hasOwnProperty('Suscripción extranjero')) {
                    for (var i in newArraySort['Suscripción extranjero']) {
                        finalObj.push(newArraySort['Suscripción extranjero'][i])
                    }
                }
                if (
                    newArraySort['Suscripción nacional'].hasOwnProperty(
                        'SEPOMEX ORDINARIO'
                    )
                ) {
                    for (var i in newArraySort['Suscripción nacional'][
                        'SEPOMEX ORDINARIO'
                    ]) {
                        finalObj.push(
                            newArraySort['Suscripción nacional'][
                            'SEPOMEX ORDINARIO'
                            ][i]
                        )
                    }
                }
                if (
                    newArraySort['Suscripción nacional'].hasOwnProperty(
                        'SEPOMEX CERTIFICADO'
                    )
                ) {
                    for (var i in newArraySort['Suscripción nacional'][
                        'SEPOMEX CERTIFICADO'
                    ]) {
                        finalObj.push(
                            newArraySort['Suscripción nacional'][
                            'SEPOMEX CERTIFICADO'
                            ][i]
                        )
                    }
                }
                if (
                    newArraySort['Suscripción nacional'].hasOwnProperty('Librerías')
                ) {
                    for (var i in newArraySort['Suscripción nacional'][
                        'Librerías'
                    ]) {
                        finalObj.push(
                            newArraySort['Suscripción nacional']['Librerías'][i]
                        )
                    }
                }
                if (
                    newArraySort['Suscripción nacional'].hasOwnProperty(
                        'SEPOMEX PAQ-POST'
                    )
                ) {
                    for (var i in newArraySort['Suscripción nacional'][
                        'SEPOMEX PAQ-POST'
                    ]) {
                        finalObj.push(
                            newArraySort['Suscripción nacional'][
                            'SEPOMEX PAQ-POST'
                            ][i]
                        )
                    }
                }
                if (
                    newArraySort['Suscripción nacional'].hasOwnProperty(
                        'DHL NORMAL NACIONAL'
                    )
                ) {
                    for (var i in newArraySort['Suscripción nacional'][
                        'DHL NORMAL NACIONAL'
                    ]) {
                        finalObj.push(
                            newArraySort['Suscripción nacional'][
                            'DHL NORMAL NACIONAL'
                            ][i]
                        )
                    }
                }
                if (
                    newArraySort['Suscripción nacional'].hasOwnProperty(
                        'FEDEX NACIONAL'
                    )
                ) {
                    for (var i in newArraySort['Suscripción nacional'][
                        'FEDEX NACIONAL'
                    ]) {
                        finalObj.push(
                            newArraySort['Suscripción nacional']['FEDEX NACIONAL'][
                            i
                            ]
                        )
                    }
                }
                if (
                    newArraySort['Suscripción nacional'].hasOwnProperty(
                        'REDPACK NACIONAL'
                    )
                ) {
                    for (var i in newArraySort['Suscripción nacional'][
                        'REDPACK NACIONAL'
                    ]) {
                        finalObj.push(
                            newArraySort['Suscripción nacional'][
                            'REDPACK NACIONAL'
                            ][i]
                        )
                    }
                }

                return finalObj
            } catch (e) {
                log.error({ title: 'Error on setSortData', details: e })
                return {}
            }
        }

        function orderData(data, key) {
            data.sort(function (a, b) {
                if (a[key] > b[key]) {
                    return 1
                }
                if (a[key] < b[key]) {
                    return -1
                }
                return 0
            })
            return data
        }

        function groupData(data, key, str) {
            var obj = {}
            if (str === false) {
                for (var dataKey in data) {
                    if (obj.hasOwnProperty(data[dataKey][key])) {
                        obj[data[dataKey][key]].push(data[dataKey])
                    } else {
                        obj[data[dataKey][key]] = []
                        obj[data[dataKey][key]].push(data[dataKey])
                    }
                }
            } else {
                for (var dataKey in data) {
                    if (obj.hasOwnProperty(data[dataKey][key].toUpperCase())) {
                        obj[data[dataKey][key].toUpperCase()].push(data[dataKey])
                    } else {
                        obj[data[dataKey][key].toUpperCase()] = []
                        obj[data[dataKey][key].toUpperCase()].push(data[dataKey])
                    }
                }
            }
            return obj
        }

        function getFileURL(fileId) {
            try {
                var fileObj = file.load({
                    id: fileId,
                })

                var fileURL = fileObj.url
                var scheme = 'https://'
                var host = url.resolveDomain({
                    hostType: url.HostType.APPLICATION,
                })

                var urlFinal = scheme + host + fileURL
                return urlFinal
            } catch (e) {
                log.error('Error on getFileURL', e)
            }
        }

        return {
            onRequest: onRequest,
        }
    })
