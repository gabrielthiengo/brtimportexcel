'use strict'

const Helpers = use('Helpers')
const Excel = use('exceljs')
const fs = use('fs')

class ImportController {
  async store({ request, response }) {

    const upload = request.file('file', { size: '5mb' })
    const fileName = `${Date.now()}.xlsx`

    await upload.move(Helpers.tmpPath('uploads'), {
      name: fileName
    })

    if(!upload.moved()) {
      return response.status(500).send({ error: 'Houve um erro ao importar' })
    }

    let workbook = new Excel.Workbook()

    workbook = await workbook.xlsx.readFile(`tmp/uploads/${fileName}`)

    const explanation = workbook.getWorksheet('Mailing')

    var excelList = []

        const colColumn = await explanation.getColumn('C')

        colColumn.eachCell((cell, index) => {
          let isPriorizado = false
          let excel = {
            solicitacao: '',
            cartao: '',
            administradora: '',
            conta: '',
            dt_abertura: '',
            nome_titular: '',
            dt_aniversario: '',
            nro_documento: '',
            nro_identificacao: '',
            email: '',
            nome_mae: '',
            nome_pai: '',
            produto: '',
            chipras: '',
            cod_bloq1: '',
            dt_bloq1: '',
            dt_bloq2: '',
            canal: '',
            matricula: '',
            proposta: '',
            acao_venda: '',
            pag_consta_pagamentos: '',
            dt_pagamento1: '',
            vl_pagamento1: '',
            dt_pagamento2: '',
            vl_pagamento2: '',
            dt_pagamento3: '',
            vl_pagamento3: '',
            dt_pagamento4: '',
            vl_pagamento4: '',
            dt_pagamento5: '',
            vl_pagamento5: '',
            dt_pagamento6: '',
            vl_pagamento6: '',
            cartao1: '',
            canal1: '',
            dt_cnt1: '',
            bp1: '',
            cartao2: '',
            canal2: '',
            dt_cnt2: '',
            bp2: '',
            cartao3: '',
            canal3: '',
            dt_cnt3: '',
            bp3: '',
            cartao4: '',
            canal4: '',
            dt_cnt4: '',
            bp4: '',
            cartao5: '',
            canal5: '',
            dt_cnt5: '',
            bp5: '',
            cartao6: '',
            canal6: '',
            dt_cnt6: '',
            bp6: '',
            vetor_end: '',
            bairro: '',
            logradouro: '',
            numero: '',
            complemento: '',
            municipio: '',
            estado: '',
            cep: '',
            city_code: '',
            pais: '',
            ddi: '',
            ddd_tel: '',
            tel: '',
            ramal: '',
            ddd_cel: '',
            celular: '',
            ddd_fax: '',
            fax: '',
            bip: '',
            ar: '',
            saldo_enquadrado: '',
            dt_inicio_fase: '',
            atividade: '',
            empresa: '',
            canal: '',
            dados_da_reclamacao: '',
            coment_reclamacao: '',
            bco_agencia: '',
            nro_cc: '',
            priorizado: false
          }

          if (index > 1) {
            excel.solicitacao = explanation.getCell('A' + index).value
            excel.cartao = explanation.getCell('B' + index).value
            excel.administradora = explanation.getCell('C' + index).value
            excel.conta = explanation.getCell('D' + index).value
            excel.dt_abertura = explanation.getCell('E' + index).value
            excel.nome_titular = explanation.getCell('F' + index).value
            excel.dt_aniversario = explanation.getCell('G' + index).value
            excel.nro_documento = explanation.getCell('H' + index).value
            excel.nro_identificacao = explanation.getCell('I' + index).value
            excel.email = explanation.getCell('J' + index).value
            excel.nome_mae = explanation.getCell('K' + index).value
            excel.nome_pai = explanation.getCell('L' + index).value
            excel.produto = explanation.getCell('M' + index).value
            excel.chipras = explanation.getCell('N' + index).value
            excel.cod_bloq1 = explanation.getCell('O' + index).value
            excel.dt_bloq1 = explanation.getCell('P' + index).value
            excel.dt_bloq2 = explanation.getCell('Q' + index).value
            excel.canal = explanation.getCell('R' + index).value
            excel.matricula = explanation.getCell('S' + index).value
            excel.proposta = explanation.getCell('T' + index).value
            excel.acao_venda = explanation.getCell('U' + index).value
            excel.pag_consta_pagamentos = explanation.getCell('V' + index).value
            excel.dt_pagamento1 = explanation.getCell('W' + index).value
            excel.vl_pagamento1 = explanation.getCell('X' + index).value
            excel.dt_pagamento2 = explanation.getCell('Y' + index).value
            excel.vl_pagamento2 = explanation.getCell('Z' + index).value
            excel.dt_pagamento3 = explanation.getCell('AA' + index).value
            excel.vl_pagamento3 = explanation.getCell('AB' + index).value
            excel.dt_pagamento4 = explanation.getCell('AC' + index).value
            excel.vl_pagamento4 = explanation.getCell('AD' + index).value
            excel.dt_pagamento5 = explanation.getCell('AE' + index).value
            excel.vl_pagamento5 = explanation.getCell('AF' + index).value
            excel.dt_pagamento6 = explanation.getCell('AG' + index).value
            excel.vl_pagamento6 = explanation.getCell('AH' + index).value
            excel.cartao1 = explanation.getCell('AI' + index).value
            excel.canal1 = explanation.getCell('AJ' + index).value
            excel.dt_cnt1 = explanation.getCell('AK' + index).value
            excel.bp1 = explanation.getCell('AL' + index).value
            excel.cartao2 = explanation.getCell('AM' + index).value
            excel.canal2 = explanation.getCell('AN' + index).value
            excel.dt_cnt2 = explanation.getCell('AO' + index).value
            excel.bp2 = explanation.getCell('AP' + index).value
            excel.cartao3 = explanation.getCell('AQ' + index).value
            excel.canal3 = explanation.getCell('AR' + index).value
            excel.dt_cnt3 = explanation.getCell('AS' + index).value
            excel.bp3 = explanation.getCell('AT' + index).value
            excel.cartao4 = explanation.getCell('AU' + index).value
            excel.canal4 = explanation.getCell('AV' + index).value
            excel.dt_cnt4 = explanation.getCell('AW' + index).value
            excel.bp4 = explanation.getCell('AX' + index).value
            excel.cartao5 = explanation.getCell('AY' + index).value
            excel.canal5 = explanation.getCell('AZ' + index).value
            excel.dt_cnt5 = explanation.getCell('BA' + index).value
            excel.bp5 = explanation.getCell('BB' + index).value
            excel.cartao6 = explanation.getCell('BC' + index).value
            excel.canal6 = explanation.getCell('BD' + index).value
            excel.dt_cnt6 = explanation.getCell('BE' + index).value
            excel.bp6 = explanation.getCell('BF' + index).value
            excel.vetor_end = explanation.getCell('BG' + index).value
            excel.bairro = explanation.getCell('BH' + index).value
            excel.logradouro = explanation.getCell('BI' + index).value
            excel.numero = explanation.getCell('BJ' + index).value
            excel.complemento = explanation.getCell('BK' + index).value
            excel.municipio = explanation.getCell('BL' + index).value
            excel.estado = explanation.getCell('BM' + index).value
            excel.cep = explanation.getCell('BN' + index).value
            excel.city_code = explanation.getCell('BO' + index).value
            excel.pais = explanation.getCell('BP' + index).value
            excel.ddi = explanation.getCell('BQ' + index).value
            excel.ddd_tel = explanation.getCell('BR' + index).value
            excel.tel = explanation.getCell('BS' + index).value
            excel.ramal = explanation.getCell('BT' + index).value
            excel.ddd_cel = explanation.getCell('BU' + index).value
            excel.celular = explanation.getCell('BV' + index).value
            excel.ddd_fax = explanation.getCell('BW' + index).value
            excel.fax = explanation.getCell('BX' + index).value
            excel.bip = explanation.getCell('BY' + index).value
            excel.ar = explanation.getCell('BZ' + index).value
            excel.saldo_enquadrado = explanation.getCell('CA' + index).value
            excel.dt_inicio_fase = explanation.getCell('CB' + index).value
            excel.atividade = explanation.getCell('CC' + index).value
            excel.empresa = explanation.getCell('CE' + index).value
            excel.canal7 = explanation.getCell('CF' + index).value
            excel.dados_da_reclamacao = explanation.getCell('CG' + index).value
            excel.coment_reclamacao = explanation.getCell('CH' + index).value
            excel.bco_agencia = explanation.getCell('CI' + index).value
            excel.nro_cc = explanation.getCell('CJ' + index).value

            const { tint } = explanation.getCell('A' + index).fill.fgColor

            if(tint) {
              isPriorizado = true
            }

            excel.priorizado = isPriorizado

            excelList.push(excel)
          }
        })

        fs.unlinkSync(`tmp/uploads/${fileName}`)

        return excelList
  }
}

module.exports = ImportController
