const xlsxFile = require('read-excel-file/node');
const nodemailer = require('nodemailer');

function sendMail(to,subject,message) 
{
   var smtpConfig = {
      service: 'Gmail',
      auth: {
          user: 'XXX@gmail.com',
          pass: 'XXX'
      }
   };
   var transporter = nodemailer.createTransport(smtpConfig);
   var mailOptions = {
      from: '"Sender Name" <sender@gmail.com>', // sender address
      to: to, // list of receivers
      subject: subject, // Subject line
      text: 'Hello world ?', // plaintext body
      html: message // html body
   };
   
   transporter.sendMail(mailOptions, function(error, info){
      if(error)
      {
         return console.log(error);
      }
      else
      {
         return console.log(info.response);
      }      
   }); 
}
var message = '<p>This is HTML content</p>';

const seeRows = () => {
   xlsxFile('./db.xlsx').then((rows) => {
      console.table(rows);
      rows.map((el, i)=> {
         // i && console.log(el, i)
         i && sendMail(el[2], el[0], `Hello ${el[1]}`)
      })
   })
}

seeRows();

// Clonar el repositorio
// Correr el comando: npm install 
// Actualizar Base de datos (db.xlsx)
// Actualizar en el archivo main.js los datos de acceso AUTH de GMAIL y datos requeridos
// Correr el comando node main.js
// Y listo! correos enviados
