const xlsxFile = require('read-excel-file/node');
const nodemailer = require('nodemailer');

function sendMail(to, subject, message) {
   var smtpConfig = {
      service: 'Gmail',
      auth: {
         user: 'XXX@gmail.com',
         pass: 'XXX'
      }
   };
   var transporter = nodemailer.createTransport(smtpConfig);
   var mailOptions = {
      from: '"Kath Montalvo" <kathmontalvo123@gmail.com>', // sender address
      to: to, // list of receivers
      subject: subject, // Subject line
      text: 'Hello world ?', // plaintext body
      html: message // html body
   };

   transporter.sendMail(mailOptions, function (error, info) {
      if (error) {
         return console.log(error);
      }
      else {
         return console.log(info.response);
      }
   });
}

const sendMailxExcel = () => {
   xlsxFile('./db.xlsx').then((rows) => {
      console.table(rows);
      rows.map((row, i) => {
         const email = row[8];
         const name = row[7];
         const lastName = row[6];
         const langEng = row[3] === 'English';
         const company = row[1];
         const valueChainPosition = row[5];
         const contacted = row[11] === 'NO';


         let fullName = '';
         if (name && lastName) { // Hay nombre y apellido
            fullName = `${name} ${lastName}`;
         } else if (name && !lastName) { // Solo nombre
            fullName = name;
         } else if (!name && lastName) { // Solo apellido
            fullName = lastName;
         } else if (!name && !lastName && langEng) { // No hay nombre ni apellido y el idioma es inglés.
            fullName = `coffe lover from ${company}`
         } else if (!name && !lastName && !langEng) { // No hay nombre ni apellido y el idioma NO es inglés (esp).
            fullName = `amante del café de ${company}`
         }

         const subject = `Value chain - ${valueChainPosition}`;

         const englishMsg = `
            Dear ${fullName}, <br/>
            I am a Peruvian student currently working on a research about specialty coffee value chain. I would like to get an approach on some key points regarding ${row[2]} and ${row[4]} markets. Would you be interested in scheduling a meeting? <br/>
            Thanks in advance!
         `;
         const spanishMsg = `
            Estimadx ${fullName}, <br/>
            Soy un estudiante peruano actualmente investigando sobre la cadena de valor de los cafés de especialidad. Me gustaría conversar contigo sobre algunos puntos claves respecto a el mercado de ${row[4]}, ${row[2]}. Crees que podamos agendar una reunión virtual?
            Muchas gracias!
         `;
         const message = langEng ? englishMsg : spanishMsg;
         i && contacted && sendMail(email, subject, message)
      })
   })
}

sendMailxExcel();

// Clonar el repositorio
// Correr el comando: npm install 
// Actualizar Base de datos (db.xlsx)
// Actualizar en el archivo main.js los datos de acceso AUTH de GMAIL y datos requeridos
// Correr el comando node main.js
// Y listo! correos enviados
