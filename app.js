const express = require('express');
const multer = require('multer');
const exceljs = require('exceljs');
const path=require('path');
const Excel = require('exceljs');
const ExcelData = require('./models/ExcelData');
const app = express();
const nodemailer = require('nodemailer');
const PDFDocument = require('pdfkit');
const ejs = require('ejs');
const pdf = require('html-pdf');

const port = 3000; // Choose any port you prefer

require('./db');

app.use(express.static(path.join(__dirname, 'public')));
app.set('view engine','ejs');
app.set('views','views');  
const storage = multer.memoryStorage(); // Store files in memory
const upload = multer({ storage: storage });
// Basic route
app.get('/', (req, res) => {
  res.render('upload');
});
// app.get('/upload', (req, res) => {
//   res.render('upload'); // Render your EJS file
// });

// Handling file upload POST request
const excelExtensions = ['.xlsx', '.xls'];


app.post('/', upload.single('excelFile'), async (req, res) => {
  // Access the uploaded file via req.file
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  // Check if the uploaded file is an Excel file (similar to previous example)
  const fileExtension = `.${req.file.originalname.split('.').pop()}`.toLowerCase();
  if (!excelExtensions.includes(fileExtension)) {
    return res.status(400).send('Please upload an Excel file (XLSX/XLS).');
  }
  // Read data from the Excel file
 
  const workbook = new exceljs.Workbook();
  try {
    await workbook.xlsx.load(req.file.buffer); // Load the uploaded file buffer into the workbook
    const worksheet = workbook.worksheets[0]; // Get the first worksheet

    const jsonData = [];
    
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber !== 1) { // Skip the header row (if applicable)
        const rowObject = {};
        row.eachCell((cell, colNumber) => {
          // Assuming the first row is headers, use the header as keys
          const headerCell = worksheet.getRow(1).getCell(colNumber);
          const header = headerCell.value;
          rowObject[header] = cell.value;
        });
        jsonData.push(rowObject);
      }
    });
    
    const insertData = jsonData.map((row) => new ExcelData(row));
    await ExcelData.insertMany(insertData);
    const recipients = [...jsonData];
    console.log(recipients);

    for (const recipient of recipients) {
      const doc = new PDFDocument();
      if(!recipient.email){
        console.log("Skipping because mail was not found!");
        continue;
      }
      const { name, course } = recipient;
      const completionDate = new Date().toDateString();
      const certificateData = {
        name,
        course,
        completionDate
      };
      ejs.renderFile(path.join(__dirname, 'views', 'certificate.ejs'), certificateData, async (err, html) => {
        if (err) {
          console.error('Error rendering EJS to HTML:', err);
          return res.status(500).send('Error rendering EJS to HTML', err);
        }
    
        // Convert HTML to PDF
        pdf.create(html).toBuffer(async (err, pdfBuffer) => {
          if (err) {
            console.error('Error converting HTML to PDF:', err);
            return res.status(500).send('Error converting HTML to PDF');
          }
          const recipientEmail = recipient.email;
          const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
              user: 'harshitarajpal1523@gmail.com',
              pass: 'rphu rkrs assw afmk',
            },
          });
          const mailOptions = {
            from: 'harshitarajpal1523@gmail.com',
            to: recipientEmail,
            subject: `Certificate of Completion of ${recipient.course} course`,
            text: `Dear ${recipient.name},
  
            Thank you for using our service. Attached is your Certificate of Completion for completing the ${recipient.course} course on ${completionDate}. We hope this certification adds value to your skills and achievements.
            
            If you have any questions or need further assistance, feel free to contact us.
            
            Best regards
            Harshita Rajpal`,
            attachments: [
              {
                filename: 'certificate.pdf',
                content: pdfBuffer,
                encoding: 'base64',
              },
            ],
          };
          // console.log(recipient);
          transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
              console.error('Error sending email:', error);
              return res.status(500).send(`Error sending email: ${error}`);
            }
            console.log('Email sent to:',recipient.name," ", info.response);
            res.send('Email sent successfully!');
          });
        });
      });

    }
  } catch (err) {
    console.error('Error reading Excel file:', err);
    res.status(500).send(`'Error reading Excel file.' ${err}`);
  }
});

// Render the certificate with data (name, course, completionDate)



app.get('/demo',(req,res)=>{
  const certificateData={
    name:"jane doe",
    course:"Web Development",
    completionDate:"1 january 1990"
  }
  res.render('certificate', certificateData);
})



// Start server
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
// app.get('/certificate', async (req, res) => {
//   try {
//     const firstRowData = await ExcelData.findOne({}).lean();
//     if (!firstRowData || !firstRowData.email) {
//       return res.status(400).send('Email address not found or empty.');
//     }
//     const doc = new PDFDocument();
//     const { name, course } = firstRowData;
//     console.log(firstRowData);
//     const completionDate=new Date().toDateString();
//     const certificateData = {
//       name,
//       course,
//       completionDate
//     };
//     ejs.renderFile(path.join(__dirname, 'views', 'certificate.ejs'), certificateData, async (err, html) => {
//       if (err) {
//         console.error('Error rendering EJS to HTML:', err);
//         return res.status(500).send('Error rendering EJS to HTML', err);
//       }
  
//       // Convert HTML to PDF
//       pdf.create(html).toBuffer(async (err, pdfBuffer) => {
//         if (err) {
//           console.error('Error converting HTML to PDF:', err);
//           return res.status(500).send('Error converting HTML to PDF');
//         }
//         const recipientEmail = firstRowData.email;
//         const transporter = nodemailer.createTransport({
//           service: 'gmail',
//           auth: {
//             user: 'harshitarajpal1523@gmail.com',
//             pass: 'rphu rkrs assw afmk',
//           },
//         });
//       // recipientEmail
//         // console.log(firstRowData.course);
//         const mailOptions = {
//           from: 'harshitarajpal1523@gmail.com',
//           to: 'harshita0592.be21@chitkara.edu.in',
//           subject: `Certificate of Completion of ${firstRowData.course} course`,
//           text: `Dear ${firstRowData.name},

//           Thank you for using our service. Attached is your Certificate of Completion for completing the ${firstRowData.course} course on ${completionDate}. We hope this certification adds value to your skills and achievements.
          
//           If you have any questions or need further assistance, feel free to contact us.
          
//           Best regards
//           Harshita Rajpal`,
//           attachments: [
//             {
//               filename: 'certificate.pdf',
//               content: pdfBuffer,
//               encoding: 'base64',
//             },
//           ],
//         };

//         transporter.sendMail(mailOptions, (error, info) => {
//           if (error) {
//             console.error('Error sending email:', error);
//             return res.status(500).send(`Error sending email: ${error}`);
//           }
//           console.log('Email sent:', info.response);
//           res.send('Email sent successfully!');
//         });
//       });
//   });
//   } catch (err) {
//     console.error('Error processing certificate and sending email:', err);
//     res.status(500).send(`Error processing certificate and sending email: ${err}`);
//   }
  
// });