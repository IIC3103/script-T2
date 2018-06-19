const excel = require('exceljs')
const req = require('request-promise')
const urls = {
    '14634759':'https://murmuring-plains-96393.herokuapp.com',
    '13634259':'https://polar-falls-30628.herokuapp.com/news',
    '14631849':'https://pacific-citadel-50177.herokuapp.com/api/v1',
    '1463225J':'https://quiet-refuge-46259.herokuapp.com',
    '12624519':'https://nameless-ravine-37937.herokuapp.com/',
    '14631814':'https://tarea2sebaguerrero.herokuapp.com/api/v1',
    '13203983':'https://peaceful-reaches-78452.herokuapp.com',
    '13638084':'https://fierce-harbor-56873.herokuapp.com',
    '13634895':'https://serene-eyrie-37109.herokuapp.com/',
    '14633388':'https://tarea1-camila-olguin.herokuapp.com/api',
    '13206028':'https://rmsolari-t2.herokuapp.com',
    '14203847':'https://desolate-garden-45879.herokuapp.com',
    '13637398':'https://hidden-caverns-16632.herokuapp.com/api/entries',
    '13636502':'https://peaceful-sea-91873.herokuapp.com/',
    '14637898':'https://apiflyingnews5.herokuapp.com',
    '13635190':'https://quiet-basin-79530.herokuapp.com/api',
    '13633376':'https://radiant-temple-77736.herokuapp.com/',
    '13636995':'https://integracion-tarea2.herokuapp.com/',
    '12634697':'https://boiling-fjord-31817.herokuapp.com/',
    '13637207':'https://apricot-cobbler-51523.herokuapp.com',
    '13633031':'https://pure-citadel-99296.herokuapp.com/',
    '15209601':'http://t2-taller-integracion-rorpis.herokuapp.com/',
    '13633732':'https://warm-citadel-87191.herokuapp.com/',
    '13635050':'https://protected-thicket-16181.herokuapp.com',
    '13200070':'https://cherry-sundae-70024.herokuapp.com',
    '13637819':'https://afternoon-garden-71964.herokuapp.com/api/v1',
    '13633309':'https://tarea-2.herokuapp.com/',
    '12634328':'https://still-depths-63230.herokuapp.com/',
    '11636343':'https://blooming-mesa-84033.herokuapp.com/',
    '13634275':'https://t2jpschelel.herokuapp.com',
    '13634542':'https://guarded-anchorage-12890.herokuapp.com/api/v1/entries',
    '13633082':'https://fake-news-api.herokuapp.com/',
    '13637851':'https://peaceful-thicket-57221.herokuapp.com/',
    '11635398':'https://limitless-tundra-22593.herokuapp.com/api/v1',
    '14638002':'https://mighty-wave-27762.herokuapp.com/api',
    '14209519':'https://protected-chamber-30459.herokuapp.com/',
    '13634836':'https://shielded-falls-77295.herokuapp.com',
    '13632906':'https://glacial-sands-98908.herokuapp.com',
    '12638536':'https://blooming-spire-69912.herokuapp.com/api/v1/entries',
}

let resultExcel = new excel.Workbook()
resultExcel.creator = 'Felipe Andres Rojos Almuna';
resultExcel.lastModifiedBy = 'Felipe Andres Rojos Almuna';
resultExcel.created = new Date();
resultExcel.modified = new Date();
resultExcel.lastPrinted = new Date();
//View options
resultExcel.views = [
    {
      x: 0, y: 0, width: 10000, height: 20000,
      firstSheet: 0, activeTab: 1, visibility: 'visible'
    }
  ]

let pointsSheet  = resultExcel.addWorksheet('Puntos');
let requestSheet = resultExcel.addWorksheet('Requests');
//Header table
pointsSheet.getCell('A1').value = 'Numero de alumno'
pointsSheet.getCell('B1').value = 'URL de la API implementada'

pointsSheet.getCell('C1').value = 'Noticia - Obtener Todas las noticias - Test formato'             
pointsSheet.getCell('D1').value = 'Noticia - Obtener Todas las noticias - Test status code'         

pointsSheet.getCell('E1').value = 'Noticia - Crear nueva noticia - Test formato'                    
pointsSheet.getCell('F1').value = 'Noticia - Crear nueva noticia - Test status code'
pointsSheet.getCell('G1').value = 'Noticia - Crear nueva noticia - Test creacion'
pointsSheet.getCell('H1').value = 'Noticia - Crear nueva noticia - Test header location'

pointsSheet.getCell('I1').value = 'Noticia - Obtener una noticia - Test formato'
pointsSheet.getCell('J1').value = 'Noticia - Obtener una noticia - Test status code'

pointsSheet.getCell('K1').value = 'Noticia - Actualizar una noticia parcialmente - Test formato'
pointsSheet.getCell('L1').value = 'Noticia - Actualizar una noticia parcialmente - Test status code'
pointsSheet.getCell('M1').value = 'Noticia - Actualizar una noticia parcialmente - Test actualizacion'

pointsSheet.getCell('N1').value = 'Noticia - Borrar una noticia - Test status code'

pointsSheet.getCell('O1').value = 'Noticia - Obtener una noticia inexistente- Test formato'
pointsSheet.getCell('P1').value = 'Noticia - Obtener una noticia inexistente- Test status code'
pointsSheet.getCell('Q1').value = 'Noticia - Obtener una noticia inexistente- Test mensaje'

pointsSheet.getCell('R1').value = 'Comentario - Obtener los comentarios de una noticia - Test formato'
pointsSheet.getCell('S1').value = 'Comentario - Obtener los comentarios de una noticia - Test status code'

pointsSheet.getCell('T1').value = 'Comentario - Crear un nueva comentario - Test formato'
pointsSheet.getCell('U1').value = 'Comentario - Crear un nueva comentario - Test status code'
pointsSheet.getCell('V1').value = 'Comentario - Crear un nueva comentario - Test creacion'
pointsSheet.getCell('W1').value = 'Comentario - Crear un nueva comentario - Test header location'

requestSheet.getCell('A1').value = 'Numero de alumno'
requestSheet.getCell('B1').value = 'Numero de request'
requestSheet.getCell('C1').value = 'Request'
requestSheet.getCell('D1').value = 'Response Header'
requestSheet.getCell('E1').value = 'Response Body'
requestSheet.getCell('F1').value = 'Response Status Code'

const formatRequestArray = (body, squema, limit=false)=>{
    body.forEach((entry)=>{
        const keys = Object.keys(entry)
        squema.forEach(s => { if(!keys.includes(s))return 0 })
        if (limit && entry["body"] && entry["body"].length > 500) return 0
    })
    return 1
}

const formatRequestSimple = (body, squema)=>{
    const keys = Object.keys(body)
    squema.forEach(s =>{ if(!keys.includes(s)) return 0 })
    return 1
}

const recoder = (bodies, headers, statuses, options, header, body, status, option)=>{
    bodies.push(body)
    headers.push(JSON.stringify(header))
    statuses.push(status)
    options.push(JSON.stringify(option))
}
const solver = async () => {
    const requests = Object.keys(urls).map( async (key) =>{
        const option1 = {
            headers: {'content-type' : 'application/json'},
            uri:     urls[key]+'/news',
            resolveWithFullResponse: true
        }
        const squema1 = [
            "id",
            "title",
            "subtitle",
            "body",
            "created_at"
        ]
        const squema2 = squema1
        const squema3 = squema1
        const squema4 = squema1
        //const squema5 = squema1
        const squema6 = [
            'error',
        ]   
        const squema7 = [
            'author',
            'comment',
            'id',
            'created_at',
        ] 
        let r1t1 = 1
        let r1t2 = 1
        let r2t1 = 1
        let r2t2 = 1
        let r2t3 = 1
        let r2t4 = 0
        let r3t1 = 1
        let r3t2 = 1
        let r4t1 = 1
        let r4t2 = 1
        let r4t3 = 0
        let r5t1 = 1
        let r6t1 = 1
        let r6t2 = 1
        let r6t3 = 0
        let r7t1 = 1
        let r7t2 = 1
        let r8t1 = 1
        let r8t2 = 1
        let r8t3 = 1
        let r8t4 = 0     
        let req1, req2, req3, req4, req41, req5, req51 ,req6, req7, req71, req72
        const headers = [] 
        const bodies  = []
        const statuses = [] 
        const options = []
        try{
            req1 = await req.get(option1)
            r1t1 = formatRequestArray(JSON.parse(req1.body), squema1, true)
            if(req1.statusCode !== 200) r1t2 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                req1.headers,
                req1.body, 
                req1.statusCode,
                option1)
        }catch(e){
            r1t1 = 0
            r1t2 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option1)
        }
        const option2 = {
            headers: {'content-type' : 'application/json'},
            uri:     urls[key]+'/news',
            resolveWithFullResponse: true,
            body:{
                'title': 'Título de la noticia creada. Testing nodejs',
                'subtitle': 'Subtítulo de la noticia creada. Testing nodejs',
                'body': 'Cuerpo de la noticia creada. Testing nodejs',
            },
            json: true
        }
        try{
            req2 = await req.post(option2)
            r2t1 = formatRequestSimple(req2.body,squema2)
            if(req2.statusCode !== 201) r2t2 = 0
            if(req2.body["title"]!== "Título de la noticia creada. Testing nodejs") r2t3 = 0
            if(req2.body["subtitle"]!== "Subtítulo de la noticia creada. Testing nodejs") r2t3 = 0
            if(req2.body["body"]!== "Cuerpo de la noticia creada. Testing nodejs") r2t3 = 0
            const reg = /(^\/news\/\d)/i
            if(
                (req2.headers['Location'] && !req2.headers['Location'].match(reg))||  
                (req2.headers['location'] && !req2.headers['location'].match(reg))) r2t4 = 1
            recoder(
                bodies,
                headers,
                statuses,
                options,
                req2.headers,
                req2.body, 
                req2.statusCode,
                option2)

        }catch(e){
            r2t1 = 0
            r2t2 = 0
            r2t3 = 0
            r2t4 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option2)

        }
        let uriSpecificNews =  urls[key]+(req2.headers['Location']|| req2.headers['location'] || 'force-error-when-no-location')
        const option3 = {
            headers: {'content-type' : 'application/json'},
            uri:     uriSpecificNews,
            resolveWithFullResponse: true,
        }
        try{
            req3 = await req.get(option3)
            r3t1 = formatRequestSimple(JSON.parse(req3.body),squema3)
            if(req3.statusCode !== 200) r3t2 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                req3.headers,
                req3.body, 
                req3.statusCode, 
                option3)

        } catch(e) {
            r3t1 = 0
            r3t2 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option3)

        }
        const option4 = {
            headers: {'content-type' : 'application/json'},
            uri:    uriSpecificNews,
            resolveWithFullResponse: true,
            body:{
                "body": "Nuevo cuerpo",
            },
            json: true
        }
        try{
            req4 = await req.patch(option4)
            r4t1 = formatRequestSimple(JSON.parse(req4.body),squema4)
            if(req4.statusCode !== 200) r4t2 = 0 
            recoder(
                bodies,
                headers,
                statuses,
                options,
                req4.headers,
                req4.body, 
                req4.statusCode,
                option4)

        } catch(e) {
            r4t1 = 0
            r4t2 = 0 
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option4)

        }
        const option41 = option3
        try{
            req41 = await req.get(option41)
            if(JSON.parse(req41.body)["body"] && JSON.parse(req41.body)["body"] === "Nuevo cuerpo") r4t3 = 1
            recoder(
                bodies,
                headers,
                statuses, 
                options,
                req41.headers,
                req41.body, 
                req41.statusCode,
                option41)
        } catch(e) {
            r4t3 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option41)
        }
        const option5 = option3
        try{
            req5 = await req.del(option5)
            if(req5.statusCode !== 200) r5t1 = 0 
            recoder(
                bodies,
                headers,
                statuses,
                options,
                req5.headers,
                req5.body, 
                req5.statusCode,
                option5)

        } catch(e) {
            r5t1 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option5)
        }
        const option51 = option3
        try{
            req51 = await req.get(option51)
            if(JSON.parse(req51.body)["body"]) r5t3 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                req51.headers,
                req51.body, 
                req51.statusCode,
                option51)
        } catch(e) {
            r5t3 = 0
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option51)
        }
        const option6 = {
            headers: {'content-type' : 'application/json'},
            uri:    urls[key]+'/news/'+'24367879',
            resolveWithFullResponse: true
        }
        try{
            req6 = await req.get(option6)
            r6t1 = formatRequestSimple(JSON.parse(req6.body),squema6)
            r6t2 = 0
            if(r6t1===1 && JSON.parse(req6.body)["error"].toLowerCase().trim()==='not found') r6t3 = 1
            recoder(
                bodies,
                headers,
                statuses,
                options,
                req6.headers,
                req6.body, 
                req6.statusCode,
                option6)
        } catch (e) {
            r6t1 = formatRequestSimple(JSON.parse(e.response.body),squema6)
            if(e.statusCode !== 404) r6t2 = 0 
            if(r6t1===1 && JSON.parse(e.response.body)["error"].toLowerCase().trim()==='not found') r6t3 = 1
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option6)
        }
        
        const option7 = option2
        try{
            req7 = await req.post(option7)
            recoder(
                bodies,
                headers,
                statuses,
                options,
                req7.headers,
                req7.body, 
                req7.statusCode,
                option7)
        } catch(e) {
            console.log('No se puede crear noticia')
            r7t1 = 0
            r7t2 = 0
            r8t1 = 0
            r8t2 = 0
            r8t3 = 0
            r8t4 = 0   
            recoder(
                bodies,
                headers,
                statuses,
                options,
                e.response.headers,
                e.response.body, 
                e.statusCode,
                option7) 
        }
        if(req7){
            const uriSpecificNewsComment = urls[key]+(req7.headers['Location']||req7.headers['location']||'forceError')+'/comments'
            const option71 = {
                headers: {'content-type' : 'application/json'},
                uri: uriSpecificNewsComment,
                resolveWithFullResponse: true,
                body:{
                    "author": "Juan Pérez",
                    "comment": "Regalando manos gratis!!",
                },
                json: true
            }
            try{
                req71 = await req.post(option71)
                await req.post(option71)
                await req.post(option71)
                await req.post(option71)
                await req.post(option71)
                recoder(
                    bodies,
                    headers,
                    statuses,
                    options,
                    req71.headers,
                    req71.body, 
                    req71.statusCode,
                    option71)
            } catch (e){
                console.log('No se puede crear comentarios')
                r7t1 = 0
                r7t2 = 0
                r8t1 = 0
                r8t2 = 0
                r8t3 = 0
                r8t4 = 0
                recoder(
                    bodies,
                    headers,
                    statuses,
                    options,
                    e.response.headers,
                    e.response.body, 
                    e.statusCode,
                    option71)
            }
            if(req71){
                const option72 = {
                    headers: {'content-type' : 'application/json'},
                    uri: uriSpecificNewsComment,
                    resolveWithFullResponse: true,
                }
                r8t1 = formatRequestSimple(JSON.parse(req71.body),squema7)
                if(req71.statusCode !== 201) r8t2 = 0 
                if(JSON.parse(req71.body)["author"]!== "Juan Pérez") r8t3 = 0
                if(JSON.parse(req71.body)["comment"]!== "Regalando manos gratis!!") r8t3 = 0
                const reg2 = /(^\/news\/\d\/comments\/\d)/i
                if(
                    (req71.headers['Location'] &&  req71.headers['Location'].match(reg2))||
                    (req71.headers['location'] &&  req71.headers['location'].match(reg2))) r8t4 = 1
                try{
                    req72 = await req.get(option72)
                    r7t1 = formatRequestArray(JSON.parse(req72.body),squema7)
                    if(req72.statusCode !== 200) r7t2 = 0 
                    recoder(
                        bodies,
                        headers,
                        statuses,
                        options,
                        req72.headers,
                        req72.body, 
                        req72.statusCode,
                        option72)
                } catch(e){
                    r7t1 = 0
                    r7t2 = 0 
                    recoder(
                        bodies,
                        headers,
                        statuses,
                        options,
                        e.response.headers,
                        e.response.body, 
                        e.statusCode,
                        option72)

                }
               
            }
        }
        return [
            key,
            urls[key],
            [
                r1t1,
                r1t2,
                r2t1,
                r2t2,
                r2t3,
                r2t4,
                r3t1,
                r3t2,
                r4t1,
                r4t2,
                r4t3,
                r5t1,
                r6t1,
                r6t2,
                r6t3,
                r7t1,
                r7t2,
                r8t1,
                r8t2,
                r8t3,
                r8t4
            ],
            [
                options,
                headers,
                bodies,
                statuses,
                
            ]]
    })
    const resolvedPromises = await Promise.all(requests)
    let requestIndex = 1
    for(let i = 0; i<resolvedPromises.length; i++){
        const instance = resolvedPromises[i]
        const index = i+2
        //Header table
        pointsSheet.getCell('A'+index).value = instance[0]
        pointsSheet.getCell('B'+index).value = instance[1]
        pointsSheet.getCell('C'+index).value = instance[2][0]             
        pointsSheet.getCell('D'+index).value = instance[2][1]
        pointsSheet.getCell('E'+index).value = instance[2][2]                    
        pointsSheet.getCell('F'+index).value = instance[2][3]
        pointsSheet.getCell('G'+index).value = instance[2][4]
        pointsSheet.getCell('H'+index).value = instance[2][5]
        pointsSheet.getCell('I'+index).value = instance[2][6]
        pointsSheet.getCell('J'+index).value = instance[2][7]
        pointsSheet.getCell('K'+index).value = instance[2][8]
        pointsSheet.getCell('L'+index).value = instance[2][9] 
        pointsSheet.getCell('M'+index).value = instance[2][10] 
        pointsSheet.getCell('N'+index).value = instance[2][11]
        pointsSheet.getCell('O'+index).value = instance[2][12]
        pointsSheet.getCell('P'+index).value = instance[2][13]
        pointsSheet.getCell('Q'+index).value = instance[2][14]
        pointsSheet.getCell('R'+index).value = instance[2][15]
        pointsSheet.getCell('S'+index).value = instance[2][16]
        pointsSheet.getCell('T'+index).value = instance[2][17]
        pointsSheet.getCell('U'+index).value = instance[2][18]
        pointsSheet.getCell('V'+index).value = instance[2][19] 
        pointsSheet.getCell('W'+index).value = instance[2][20]
        for(let j = 0; j<instance[3][0].length; j++){
            requestIndex+=1
            requestSheet.getCell('A'+requestIndex).value = instance[0]
            requestSheet.getCell('B'+requestIndex).value = j+1
            requestSheet.getCell('C'+requestIndex).value = instance[3][0][j]
            requestSheet.getCell('D'+requestIndex).value = instance[3][1][j]
            requestSheet.getCell('E'+requestIndex).value = instance[3][2][j]
            requestSheet.getCell('F'+requestIndex).value = instance[3][3][j]
        }
    }
    resultExcel.xlsx.writeFile('results.xlsx').then(()=>{
        console.log('Tareas revisadas :)')
    }).catch((e)=>{
        console.error('No se pueden crea excel')
    })
    return resolvedPromises
    //const result = await resultExcel.xlsx.writeFile('results.xlsx')
    //console.log('Archivo guardado')
}
const res = solver().then(result => result)

/*resultExcel.xlsx.writeFile('results.xlsx')
    .then(function() {
        //promise
        console.log('Archivo guardado')
    });*/
