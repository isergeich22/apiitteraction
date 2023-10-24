const http = require('http')
const express = require('express')
const fetch = require('node-fetch')
const components = require('./components.js')
const exl = require('exceljs')
const fs = require('fs')
const cio = require('cheerio')

const app = express()

const dbsId = '26502764'
const fbsId = '70747147'

app.get(['/', 'home'], async function(req, res){

    //fullfType=fbs {
        
    const natCatPath = './public/Краткий отчет.xlsx'

    let natCatList = []
    let notMarkedOrders = []
    let products = []
    let productsArticles = []

    let productsList = []
    let productsArticlesList = []

    async function getOrders(clientId) {

        let response = await fetch(`https://api.partner.market.yandex.ru/campaigns/${clientId}/orders?status=PROCESSING&substatus=STARTED`, {
            method: 'GET',
            headers: {
                'Authorization': 'Bearer y0_AgAAAAAicc0tAAqYtAAAAADuPsgggQLpqr8rScK5QFQyQQqjRimPEkI'
            }
        })

        let result = await response.json()

        result.orders.forEach(elem => {

            // console.log(elem)

            elem.items.forEach(el => {
                
                if(el.requiredInstanceTypes) {
                    if(el.requiredInstanceTypes.indexOf('CIS') >= 0) {

                        if(el.instances === undefined) {

                            products.push(el.offerName)
                            productsArticles.push(el.offerId)
                            notMarkedOrders.push(elem)

                        }

                    }
                }

            })
        })

    }

    await getOrders(dbsId)

    let html = ``

    async function getProductsList() {

        let items = []

        for(let i = 0; i < products.length; i++) {
    
            if(products[i].indexOf('Maktex') >= 0) {
    
                let response = await fetch('https://api-seller.ozon.ru/v2/product/info', {
                    method: 'POST',
                    headers: {
                        'Host': 'api-seller.ozon.ru',
                        'Client-Id': '144225',
                        'Api-Key': '5d5a7191-2143-4a65-ba3a-b184958af6e8',
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        "offer_id": `${productsArticles[i]}`,
                        "product_id": 0,
                        "sku": 0
                    })
                })
    
                let result = await response.json()
    
                // html += `<p>${result.result.name} - <span>${productsArticles[i]}</span></p>`
    
                items.push(result.result.name.trim())
                productsArticlesList.push(productsArticles[i])
    
            }
    
            if(products[i].indexOf('Maktex') < 0) {
    
                // html += `<p>${products[i]} - <span>${productsArticles[i]}</span></p>`
    
                items.push(products[i].trim())
                productsArticlesList.push(productsArticles[i])
    
            }            
    
        }

        return items

    }    

    productsList = await getProductsList()    

    async function getNationalCatalog(filepath) {

        let nat_cat = []

        const wb = new exl.Workbook()

        await wb.xlsx.readFile(filepath)

        const ws = wb.getWorksheet('Краткий отчет')

            // console.log(ws)

        const c2 = ws.getColumn(2)

        // console.log(c2)

        c2.eachCell(c => {
            // console.log(c.value)
            nat_cat.push(c.value)
        })

        // console.log(nat_cat)
        return nat_cat

    }

    // console.log(productsList)

    natCatList = await getNationalCatalog(natCatPath)

    // console.log(natCatList)

    let newProductsList = []

    async function renderList() {

        productsList.forEach(el => {

            if(natCatList.indexOf(el) < 0) {

                html += `<p><span>Новый</span> - ${el} - <span>${productsArticlesList[productsList.indexOf(el)]}</span></p>`
                newProductsList.push(el)

            }

            if(natCatList.indexOf(el) >= 0) {

                html += `<p><span>Актуальный</span> - ${el} - <span>${productsArticlesList[productsList.indexOf(el)]}</span></p>`

            }

        })

    }
    
    await renderList()

    async function createImport(new_products) {

        const filepath = './public/IMPORT_TNVED_6302 (3).xlsx'

        const wb = new exl.Workbook()

        await wb.xlsx.readFile(filepath)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let startCellNumber = 5

        for(let i = 0; i < new_products.length; i++) {

            let size = ''
            ws.getCell(`A${startCellNumber}`).value = 6302
            ws.getCell(`B${startCellNumber}`).value = new_products[i]
            ws.getCell(`C${startCellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`D${startCellNumber}`).value = 'Артикул'
            ws.getCell(`E${startCellNumber}`).value = productsArticlesList[productsList.indexOf(new_products[i])]
            ws.getCell(`H${startCellNumber}`).value = 'ВЗРОСЛЫЙ'
            if(new_products[i].indexOf('Постельное') >= 0 || new_products[i].indexOf('Детское') >= 0) {
                ws.getCell(`F${startCellNumber}`).value = 'КОМПЛЕКТ'
            }

            if(new_products[i].indexOf('Полотенце') >= 0) {
                ws.getCell(`F${startCellNumber}`).value = 'ИЗДЕЛИЯ ДЛЯ САУНЫ'
            }
            
            if(new_products[i].indexOf('Простыня') >= 0) {
                if(new_products[i].indexOf('на резинке') >= 0) {
                    ws.getCell(`F${startCellNumber}`).value = 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                } else {
                    ws.getCell(`F${startCellNumber}`).value = 'ПРОСТЫНЯ'
                }
            }
            if(new_products[i].indexOf('Пододеяльник') >= 0) {
                ws.getCell(`F${startCellNumber}`).value = 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
            }
            if(new_products[i].indexOf('Наволочка') >= 0) {
                if(new_products[i].indexOf('50х70') >=0 || new_products[i].indexOf('40х60') >= 0 || new_products[i].indexOf('50 х 70') >=0 || new_products[i].indexOf('40 х 60') >= 0) {
                    ws.getCell(`F${startCellNumber}`).value = 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                } else {
                    ws.getCell(`F${cellNumber}`).value = 'НАВОЛОЧКА КВАДРАТНАЯ'
                }
            }
            if(new_products[i].indexOf('Наматрасник') >= 0) {
                ws.getCell(`F${startCellNumber}`).value = 'НАМАТРАСНИК'
            }
            if(new_products[i].indexOf('страйп-сатин') >= 0 || new_products[i].indexOf('страйп сатин') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'СТРАЙП-САТИН'
            }
            if(new_products[i].indexOf('твил-сатин') >= 0 || new_products[i].indexOf('твил сатин') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'ТВИЛ-САТИН'
            }
            if(new_products[i].indexOf('тенсел') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'ТЕНСЕЛЬ'
            }
            if(new_products[i].indexOf('бяз') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'БЯЗЬ'
            }
            if(new_products[i].indexOf('поплин') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'ПОПЛИН'
            }
            if(new_products[i].indexOf('сатин') >= 0 && new_products[i].indexOf('-сатин') < 0 && new_products[i].indexOf('п сатин') < 0 && new_products[i].indexOf('л сатин') < 0 && new_products[i].indexOf('сатин-') < 0 && new_products[i].indexOf('сатин ж') < 0) {
                ws.getCell(`I${startCellNumber}`).value = 'САТИН'
            }
            if(new_products[i].indexOf('вареный') >= 0 || new_products[i].indexOf('варёный') >= 0 || new_products[i].indexOf('вареного') >= 0 || new_products[i].indexOf('варёного') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            }
            if(new_products[i].indexOf('сатин-жаккард') >= 0 || new_products[i].indexOf('сатин жаккард') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'САТИН-ЖАККАРД'
            }
            if(new_products[i].indexOf('страйп-микрофибр') >= 0 || new_products[i].indexOf('страйп микрофибр') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'МИКРОФИБРА'
            }
            if(new_products[i].indexOf('шерст') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'ПОЛИЭФИР'
            }
            if(new_products[i].indexOf('перкал') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'ПЕРКАЛЬ'
            }
            if(new_products[i].indexOf('махра') >= 0 || new_products[i].indexOf('махровое') >= 0) {
                ws.getCell(`I${startCellNumber}`).value = 'МАХРОВАЯ ТКАНЬ'
            }

            if(new_products[i].indexOf('тенсел') >= 0) {ws.getCell(`J${startCellNumber}`).value = '100% Эвкалипт'}
            else if(new_products[i].indexOf('шерст') >= 0) {ws.getCell(`J${startCellNumber}`).value = '100% Полиэстер'}
            else {ws.getCell(`J${startCellNumber}`).value = '100% Хлопок'}

            if(new_products[i].indexOf('Постельное') >= 0) {
                if(new_products[i].indexOf('1,5 спальное') >= 0 || new_products[i].indexOf('1,5 спальный') >= 0) {
                    size = '1,5 спальное'
                    if(new_products[i].indexOf('на резинке') >= 0) {
                        size += ' на резинке'
                        for(let k = 40; k < 305; k+=5) {
                            for(let l = 40; l < 305; l+=5) {
                                if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                    for(let j = 10; j < 50; j+=10) {
                                        if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                            size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                            ws.getCell(`K${startCellNumber}`).value = size
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                        size += ' с наволочками 50х70'
                        ws.getCell(`K${startCellNumber}`).value = size
                    } else {
                        ws.getCell(`K${startCellNumber}`).value = size
                    }
                }
                if(new_products[i].indexOf('2 спальное') >= 0 || new_products[i].indexOf('2 спальный') >= 0) {
                    size = '2 спальное'
                    if(new_products[i].indexOf('с Евро') >= 0) {
                        size += ' с Евро простыней'
                    }
                    if(new_products[i].indexOf('на резинке') >= 0) {
                        size += ' на резинке'
                        for(let k = 40; k < 305; k+=5) {
                            for(let l = 40; l < 305; l+=5) {
                                if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                    for(let j = 10; j < 50; j+=10) {
                                        if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                            size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                            ws.getCell(`K${startCellNumber}`).value = size
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                        size += ' с наволочками 50х70'
                        ws.getCell(`K${startCellNumber}`).value = size
                    } else {
                        ws.getCell(`K${startCellNumber}`).value = size
                    }
                }
                if(new_products[i].indexOf('Евро -') >= 0 || new_products[i].indexOf('евро -') >= 0 || new_products[i].indexOf('Евро на') >= 0 || new_products[i].indexOf('евро на') >= 0) {
                    size = 'Евро'
                    if(new_products[i].indexOf('на резинке') >= 0) {
                        size += ' на резинке'
                        for(let k = 40; k < 305; k+=5) {
                            for(let l = 40; l < 305; l+=5) {
                                if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                    for(let j = 10; j < 50; j+=10) {
                                        if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                            size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                            ws.getCell(`K${startCellNumber}`).value = size
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                        size += ' с наволочками 50х70'
                        ws.getCell(`K${startCellNumber}`).value = size
                    } else {
                        ws.getCell(`K${startCellNumber}`).value = size
                    }
                }
                if(new_products[i].indexOf('Евро Макси') >= 0 || new_products[i].indexOf('евро макси') >= 0 || new_products[i].indexOf('Евро макси') >= 0) {
                    size = 'Евро Макси'
                    if(new_products[i].indexOf('на резинке') >= 0) {
                        size += ' на резинке'
                        for(let k = 40; k < 305; k+=5) {
                            for(let l = 40; l < 305; l+=5) {
                                if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                    for(let j = 10; j < 50; j+=10) {
                                        if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                            size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                            ws.getCell(`K${startCellNumber}`).value = size
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                        size += ' с наволочками 50х70'
                        ws.getCell(`K${startCellNumber}`).value = size
                    } else {
                        ws.getCell(`K${startCellNumber}`).value = size
                    }
                }
                if(new_products[i].indexOf('семейное') >= 0 || new_products[i].indexOf('семейный') >= 0) {
                    size = 'семейное'
                    if(new_products[i].indexOf('на резинке') >= 0) {
                        size += ' на резинке'
                        for(let k = 40; k < 305; k+=5) {
                            for(let l = 40; l < 305; l+=5) {
                                if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                    for(let j = 10; j < 50; j+=10) {
                                        if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                            size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                            ws.getCell(`K${startCellNumber}`).value = size
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                        size += ' с наволочками 50х70'
                        ws.getCell(`K${startCellNumber}`).value = size
                    } else {
                        ws.getCell(`K${startCellNumber}`).value = size
                    }
                }
            } else {
                for(let k = 40; k < 305; k+=5) {
                    for(let l = 40; l < 305; l+=5) {
                        if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                            size = `${k.toString()}х${l.toString()}`
                            for(let j = 10; j < 50; j+=10) {
                                if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                    size = `${k.toString()}х${l.toString()}х${j.toString()}`
                                    ws.getCell(`K${startCellNumber}`).value = size
                                } else {
                                    ws.getCell(`K${startCellNumber}`).value = size
                                }
                            }
                        }
                    }
                }
            }

            ws.getCell(`L${startCellNumber}`).value = '6302100001'
            ws.getCell(`M${startCellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
            ws.getCell(`N${startCellNumber}`).value = 'На модерации'

            startCellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/yandex/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_yandex` : filePath = `./public/yandex/IMPORT_TNVED_6302_${date_ob.getDate()}_${month}_yandex`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    if(newProductsList.length > 0) {

        await createImport(newProductsList)

    }    

    res.send(html)

    // }

})

//     let response = await fetch('https://api.partner.market.yandex.ru/campaigns', {
//         method: 'GET',
//         headers: {
//             'Authorization': 'Bearer y0_AgAAAAAicc0tAAqYtAAAAADuPsgggQLpqr8rScK5QFQyQQqjRimPEkI'
//         }
//     })

//     let result = await response.json()

//     console.log(result)

//     res.send(`${result}`)


    
// })

app.listen(3000)