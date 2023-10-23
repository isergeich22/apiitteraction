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

    let notMarkedOrders = []
    let products = []

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
                            notMarkedOrders.push(elem)

                        }

                    }
                }

            })
        })

    }

    await getOrders(fbsId)

    

    // console.log(notMarkedOrders)

    res.send(`${products}`)

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