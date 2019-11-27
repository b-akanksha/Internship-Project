const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const ejs = require('ejs');
const GoogleSpreadsheet = require('google-spreadsheet');
const mongoose = require('mongoose');
const schema = mongoose.Schema;
const session = require('express-session');
const cookieParser = require('cookie-parser');
const { promisify } = require('util');

const creds = require('./client_secret.json');

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({
    extended: false
}));

app.use(cookieParser());
app.use(session({
    secret: "Secret123",
}));

mongoose.connect('mongodb+srv://akira:akira6@cluster0-xbelx.mongodb.net/test?retryWrites=true&w=majority', {
    useNewUrlParser: true}, 
    error => console.error(error));

app.use('/public', express.static(path.join(__dirname, "./public")));

app.set('view engine', 'ejs');
app.set('views', __dirname);

const User = new schema({
    email:String,
    name: String,
    password: String
});

const userModel = mongoose.model("user", User);

//getting faculty details
var img = [], fName = [], pos = [];
async function facDetails() {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet = info.worksheets[12];
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        img.push(row.link);
        fName.push(row.name);
        pos.push(row.position);
    })
}

var fac = {img:img, name:fName, pos:pos};

//to call index.ejs
app.get('/', (request, response) => {
    facDetails();
    response.render("index", {
        faculty:fac
    });
});

//to open about.ejs
app.get('/about', (request, response) => {
    response.render("about");
});

//to open login page
app.get('/faculty_login', (request, response) => {
    response.render("faculty_login", {
        errors: [],
    });
});

//to login page and render faculty.ejs
app.post('/faculty_login',(request,response) => {
    if(request.body.email && request.body.password) {
        if(request.body.password.trim().length > 3) {
            userModel.findOne({email: request.body.email, password: request.body.password}, (err, result) => {
                if(err) {
                    console.log(err);
                }
                if(result) {
                    request.session.user = result;
                    console.log(result);
                    response.redirect('/faculty');      
                } else {
                    response.render('faculty_login', {
                        errors: ['Invalid email id - password combination.']
                    })
                }
            })
        } else {
            response.render('faculty_login', {
                errors: ['Password should be of more characters.']
            });
        }
    } else {
        const error = [];
        if( !request.body.email) {
            error.push("Enter your email address.");
        }
        if(!request.body.password) {
            error.push("Enter your password.");
        }
        response.render('faculty_login', {
            errors: error
        });
    }
});

//to maintain session
function sessionCheck(req, res, next) {
    if(req.session.user) {
        next();
    } else {
        res.redirect('/faculty_login');
    }
}

//to open faculty
app.get('/faculty', sessionCheck, (request, response) => {
    accessSpreadsheet();
    response.render("faculty", {
        usn: usnall,
        user: request.session.user
    });
});
 
//to logout and end the session
app.get('/logout', (request, response) => {
    request.session.destroy();
    response.redirect('/');
});

//all the variables used
var usn5 = [], usn7 = [], name5 = [], sec = [], name7 = [], sub51 = [], sub52 = [], sub53 = [], sub54 = [], subele51 = [], subele52 = [], sublab51 = [], sublab52 = [], sub71 = [], sub72 = [], sub73 = [], subele71 = [], subele72 = [], sublab71 = [], sublab72 = [], high5 = [], high7 = [], avg5 = [], avg7 = [], small5 = [], small7 = [], avg5a = [], avg5b = [], avg7a = [], avg7b = [], usn510 = [], usn710 = [], name510 = [], name710 = [], usnall = [], studName, link5a = [], link5b = [], link7a = [], link7b = [], name5a = [], name7a = [], name5b = [], name7b = [], usn5a = [], usn5b = [], usn7a = [], usn7b = [];

//to access the spreadsheet and find various values
async function accessSpreadsheet() {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    const sheet = info.worksheets[3];
    const sheet1 = info.worksheets[7];
    sheet2 = info.worksheets[8];
    sheet3 = info.worksheets[9];
    console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);

    const r510 = await promisify(sheet2.getRows)({
        offset: 1,
        limit: 10
    })
    r510.forEach(row => {
        usn510.push(row.usn);
        name510.push(row.name)
    })
    const r710 = await promisify(sheet3.getRows)({
        offset: 1,
        limit: 10
    })
    r710.forEach(row => {
        usn710.push(row.usn);
        name710.push(row.name)
    })

    //To read data from sheet
    const rows = await promisify(sheet.getRows)({
        offset: 6,
    });
    for (var i = 0; i < 30; i++) {
        if (i < 15) {
            sec.push("A");
        }
        else sec.push("B");
    }
    rows.forEach(row => {
        usn5.push(row.usn);
        name5.push(row.name);
        if(row.section == 'A'){
            name5a.push(row.name);
        }else {
            name5b.push(row.name);
        }
        sub51.push(Number(row.subject1));
        sub52.push(Number(row.subject2));
        sub53.push(Number(row.subject3));
        sub54.push(Number(row.subject4));
        subele51.push(Number(row.elective1));
        subele52.push(Number(row.elective2));
        sublab51.push(Number(row.lab1));
        sublab52.push(Number(row.lab2));
    });

    const rows7 = await promisify(sheet1.getRows)({
        offset: 6,
    });;
    rows7.forEach(row => {
        usn7.push(row.usn);
        name7.push(row.name);
        if(row.section == 'A'){
            name7a.push(row.name);
        }else {
            name7b.push(row.name);
        }
        sub71.push(Number(row.subject1));
        sub72.push(Number(row.subject2));
        sub73.push(Number(row.subject3));
        subele71.push(Number(row.elective1));
        subele72.push(Number(row.elective2));
        sublab71.push(Number(row.lab1));
        sublab72.push(Number(row.lab2));
    });

    const ro5 = await promisify(sheet.getRows)({
        offset: 1,
        limit: 5
    });
    ro5.forEach(row => {
        if (row.name == 'Highest') {
            high5.push(Number(row.subject1));
            high5.push(Number(row.subject2));
            high5.push(Number(row.subject3));
            high5.push(Number(row.subject4));
            high5.push(Number(row.elective1));
            high5.push(Number(row.elective2));
            high5.push(Number(row.lab1));
            high5.push(Number(row.lab2));
        }
        if (row.name == 'Average') {
            avg5.push(Number(row.subject1));
            avg5.push(Number(row.subject2));
            avg5.push(Number(row.subject3));
            avg5.push(Number(row.subject4));
            avg5.push(Number(row.elective1));
            avg5.push(Number(row.elective2));
            avg5.push(Number(row.lab1));
            avg5.push(Number(row.lab2));
        }
        if (row.name == 'Least') {
            small5.push(Number(row.subject1));
            small5.push(Number(row.subject2));
            small5.push(Number(row.subject3));
            small5.push(Number(row.subject4));
            small5.push(Number(row.elective1));
            small5.push(Number(row.elective2));
            small5.push(Number(row.lab1));
            small5.push(Number(row.lab2));
        }
        if (row.name == 'AverageA') {
            avg5a.push(Number(row.subject1));
            avg5a.push(Number(row.subject2));
            avg5a.push(Number(row.subject3));
            avg5a.push(Number(row.subject4));
            avg5a.push(Number(row.elective1));
            avg5a.push(Number(row.elective2));
            avg5a.push(Number(row.lab1));
            avg5a.push(Number(row.lab2));
        }
        if (row.name == 'AverageB') {
            avg5b.push(Number(row.subject1));
            avg5b.push(Number(row.subject2));
            avg5b.push(Number(row.subject3));
            avg5b.push(Number(row.subject4));
            avg5b.push(Number(row.elective1));
            avg5b.push(Number(row.elective2));
            avg5b.push(Number(row.lab1));
            avg5b.push(Number(row.lab2));
        }
    });
    const ro7 = await promisify(sheet1.getRows)({
        offset: 1,
        limit: 5
    });
    ro7.forEach(row => {
        if (row.name == 'Highest') {
            high7.push(Number(row.subject1));
            high7.push(Number(row.subject2));
            high7.push(Number(row.subject3));
            high7.push(Number(row.elective1));
            high7.push(Number(row.elective2));
            high7.push(Number(row.lab1));
            high7.push(Number(row.lab2));
        }
        if (row.name == 'Average') {
            avg7.push(Number(row.subject1));
            avg7.push(Number(row.subject2));
            avg7.push(Number(row.subject3));
            avg7.push(Number(row.elective1));
            avg7.push(Number(row.elective2));
            avg7.push(Number(row.lab1));
            avg7.push(Number(row.lab2));
        }
        if (row.name == 'Least') {
            small7.push(Number(row.subject1));
            small7.push(Number(row.subject2));
            small7.push(Number(row.subject3));
            small7.push(Number(row.elective1));
            small7.push(Number(row.elective2));
            small7.push(Number(row.lab1));
            small7.push(Number(row.lab2));
        }
        if (row.name == 'AverageA') {
            avg7a.push(Number(row.subject1));
            avg7a.push(Number(row.subject2));
            avg7a.push(Number(row.subject3));
            avg7a.push(Number(row.elective1));
            avg7a.push(Number(row.elective2));
            avg7a.push(Number(row.lab1));
            avg7a.push(Number(row.lab2));
        }
        if (row.name == 'AverageB') {
            avg7b.push(Number(row.subject1));
            avg7b.push(Number(row.subject2));
            avg7b.push(Number(row.subject3));
            avg7b.push(Number(row.elective1));
            avg7b.push(Number(row.elective2));
            avg7b.push(Number(row.lab1));
            avg7b.push(Number(row.lab2));
        }
    });
    usn5.forEach(u => {
        usnall.push(u);
    })
    usn7.forEach(u => {
        usnall.push(u);
    })
}

//json passed to various ejs files
var student5 = { sem: "5th Semester", no: 5, name: name5, usn: usn5, sect: sec, subject1: sub51, subject2: sub52, subject3: sub53, subject4: sub54, ele1: subele51, ele2: subele52, lab1: sublab51, lab2: sublab52, highest: high5, average: avg5, least: small5 };
var student7 = { sem: "7th Semester", n0: 7, name: name7, usn: usn7, sect: sec, subject1: sub71, subject2: sub72, subject3: sub73, ele1: subele71, ele2: subele72, lab1: sublab71, lab2: sublab72, highest: high7, average: avg7, least: small7 };
var stud5a = { sem: "5th Semester A Section", no: 5, average: avg5a };
var stud5b = { sem: "5th Semester B Section", no: 5, average: avg5b };
var stud7a = { sem: "7th Semester A Section", no: 7, average: avg7a };
var stud7b = { sem: "7th Semester B Section", no: 7, average: avg7b };
var studtop5 = { sem: "5th Semester", usn: usn510, name: name510 };
var studtop7 = { sem: "7th Semester", usn: usn710, name: name710 };

//to get images, usn and name of students
async function images() {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet = info.worksheets[10];
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if(row.section == 'A') {
            usn5a.push(row.usn);
            link5a.push(row.link);
            name5a.push(row.name);
        } else {
            usn5b.push(row.usn);
            link5b.push(row.link);
            name5b.push(row.name);        }
    })
    var sheet1 = info.worksheets[11];
    const rowss = await promisify(sheet1.getRows)({
        offset: 1
    })
    rowss.forEach(row => {
        if(row.section == 'A') {
            usn7a.push(row.usn);
            link7a.push(row.link);
            name7a.push(row.name);
        } else {
            usn7b.push(row.usn);
            link7b.push(row.link);
            name7b.push(row.name);        }
    })
}

//json of student details(img, name, usn)
var stud5 = {link5a: link5a, link5b: link5b, name5a: name5a, usn5a: usn5a, name5b: name5b, usn5b:usn5b};
var stud7 = {link7a:link7a, link7b:link7b,name7a: name7a, usn7a: usn7a, name7b: name7b, usn7b:usn7b};

//student.ejs rendered
app.get('/student', (request, response) => {
    accessSpreadsheet();
    images();
    response.render("student", {
        usn: usnall,
        stud5 : stud5,
        stud7: stud7
    });
});

//answer to subject wise performance
app.post('/subjectwise', sessionCheck, (request, response) => {
    accessSpreadsheet();
    sem = request.body.sub_sem;
    if (sem == 5) {
        response.render('subjectwisefac', {
            student: student5,
            user: request.session.user
        })
    }
    if (sem == 7) {
        response.render('subjectwisefac', {
            student: student7,
            user:request.session.user
        })
    }
});

//answer to section wise performance
app.post('/sectionwise', sessionCheck, (request, response) => {
    accessSpreadsheet();
    sem = request.body.sec_sem;
    sect = request.body.sec_section;
    if (sem == 5) {
        if (sect == 'A') {
            response.render('sectionwisefac', {
                student: stud5a,
                student2: stud5b,
                user: request.session.user
            })
        } else {
            response.render('sectionwisefac', {
                student: stud5b,
                student2: stud5a,
                user:request.session.user
            })
        }
    }
    else {
        if (sect == 'A') {
            response.render('sectionwisefac', {
                student: stud7a,
                student2: stud7b,
                user:request.session.user
            })
        } else {
            response.render('sectionwisefac', {
                student: stud7b,
                student2: stud7a,
                user:request.session.user
            })
        }
    }
});

//top10 students of each sem
app.post('/top10', sessionCheck, (request, response) => {
    accessSpreadsheet();
    sem = request.body.top_sem;

    if (sem == 5) {
        response.render('top10', {
            student: studtop5,
            user: request.session.user
        });
    }
    else {
        response.render('top10', {
            student: studtop7,
            user:request.session.user
        });
    }
})

//update marks of sub1
async function sub1(sem, test, usn, marks) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet;
    if (sem == 5) {
        if (test == 1) {
            sheet = info.worksheets[0];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[1];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[2];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    else {
        if (test == 1) {
            sheet = info.worksheets[4];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[5];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[6];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if (row.usn == usn) {
            row.subject1 = marks;
            row.save();
        }
    })
}

//update marks of sub2
async function sub2(sem, test, usn, marks) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet;
    if (sem == 5) {
        if (test == 1) {
            sheet = info.worksheets[0];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[1];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[2];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    else {
        if (test == 1) {
            sheet = info.worksheets[4];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[5];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[6];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if (row.usn == usn) {
            row.subject2 = marks;
            row.save();
        }
    })
}

//update marks of sub3
async function sub3(sem, test, usn, marks) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet;
    if (sem == 5) {
        if (test == 1) {
            sheet = info.worksheets[0];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[1];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[2];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    else {
        if (test == 1) {
            sheet = info.worksheets[4];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[5];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[6];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if (row.usn == usn) {
            row.subject3 = marks;
            row.save();
        }
    })
}

//update marks of sub4
async function sub4(sem, test, usn, marks) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet;
    if (sem == 5) {
        if (test == 1) {
            sheet = info.worksheets[0];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[1];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[2];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    else {
        if (test == 1) {
            sheet = info.worksheets[4];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[5];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[6];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if (row.usn == usn) {
            row.subject4 = marks;
            row.save();
        }
    })
}

//update marks of ele1
async function ele1(sem, test, usn, marks) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet;
    if (sem == 5) {
        if (test == 1) {
            sheet = info.worksheets[0];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[1];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[2];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    else {
        if (test == 1) {
            sheet = info.worksheets[4];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[5];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[6];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if (row.usn == usn) {
            row.elective1 = marks;
            row.save();
        }
    })
}

//update marks of ele2
async function ele2(sem, test, usn, marks) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet;
    if (sem == 5) {
        if (test == 1) {
            sheet = info.worksheets[0];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[1];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[2];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    else {
        if (test == 1) {
            sheet = info.worksheets[4];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[5];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[6];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if (row.usn == usn) {
            row.elective2 = marks;
            row.save();
        }
    })
}

//update marks of lab1
async function lab1(sem, test, usn, marks) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet;
    if (sem == 5) {
        if (test == 1) {
            sheet = info.worksheets[0];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[1];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[2];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    else {
        if (test == 1) {
            sheet = info.worksheets[4];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[5];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[6];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if (row.usn == usn) {
            row.lab1 = marks;
            row.save();
        }
    })
}

//update marks of lab2
async function lab2(sem, test, usn, marks) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet;
    if (sem == 5) {
        if (test == 1) {
            sheet = info.worksheets[0];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[1];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[2];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    else {
        if (test == 1) {
            sheet = info.worksheets[4];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 2) {
            sheet = info.worksheets[5];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
        if (test == 3) {
            sheet = info.worksheets[6];
            console.log(`Title: ${sheet.title}, Row: ${sheet.rowCount}, Column: ${sheet.colCount}`);
        }
    }
    const rows = await promisify(sheet.getRows)({
        offset: 1
    })
    rows.forEach(row => {
        if (row.usn == usn) {
            row.lab2 = marks;
            row.save();
        }
    })
}

//to get student.ejs
app.get('/update', sessionCheck, (request,response)=> {
    response.render('faculty', {
        user: request.session.user
    });
})

//to update marks
app.post('/update',  sessionCheck, (request, response) => {
    accessSpreadsheet();
    sem = request.body.up_sem;

    if (sem == 5) {
        response.render('updatemark', {
            usn: usn5,
            student: student5,
        })
    } else {
        response.render('updatemarks', {
            usn: usn7,
            student: student7,
        })
    }
})

//to update marks of sem5
app.post('/display', (request, response) => {
    usn = request.body.usn;
    marks = request.body.mark;
    subj = request.body.sub;
    test = request.body.test;
    if(marks > 20) {
        alert('Enter marks less than 20!');
        response.render('updatemark', {
            usn: usn5,
            student: student5,
        })    
    }
    else {
        if (subj == 'subject1') {
            sub1(5, test, usn, marks);
        }
        if (subj == 'subject2') {
            sub2(5, test, usn, marks);
        }
        if (subj == 'subject3') {
            sub3(5, test, usn, marks);
        }
        if (subj == 'subject4') {
            sub4(5, test, usn, marks);
        }
        if (subj == 'elective1') {
            ele1(5, test, usn, marks);
        }
        if (subj == 'elective2') {
            ele2(5, test, usn, marks);
        }
        if (subj == 'lab1') {
            lab1(5, test, usn, marks);
        }
        if (subj == 'lab2') {
            lab2(5, test, usn, marks);
        }
        response.render('updatemark', {
            usn: usn5,
            student: student5,
            user:request.session.user
        })
    }
})

//to update marks of sem7
app.post('/display_', (request, response) => {
    usn = request.body.up_usn;
    marks = request.body.marks;
    subj = request.body.up_sub;
    test = request.body.up_test;
    if(marks > 20) {
        alert('Enter marks less than 20!');
        
        response.render('updatemarks', {
            usn: usn7,
            student: student7,
        })
    }
    else {
        if (subj == 'subject1') {
            sub1(7, test, usn, marks);
        }
        if (subj == 'subject2') {
            sub2(7, test, usn, marks);
        }
        if (subj == 'subject3') {
            sub3(7, test, usn, marks);
        }
        if (subj == 'subject4') {
            sub4(7, test, usn, marks);
        }
        if (subj == 'elective1') {
            ele1(7, test, usn, marks);
        }
        if (subj == 'elective2') {
            ele2(7, test, usn, marks);
        }
        if (subj == 'lab1') {
            lab1(7, test, usn, marks);
        }
        if (subj == 'lab2') {
            lab2(7, test, usn, marks);
        }
        response.render('updatemarks', {
            usn: usn7,
            student: student7,
            user:request.session.user
        })
    }
})

//to get values of each subject marks of a particular student
var sub5 = [], sub7 = [];
async function analysisSem(usn,sem) {
    const doc = new GoogleSpreadsheet('1ZwulnnmZo1Dx6ha7jlc0D0_N1ccV05uPqf3xeV9JTFI');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    var sheet; 
    if(sem == 5){
        sheet = info.worksheets[3];
        const rows = await promisify(sheet.getRows)({
            offset: 6
        })
        rows.forEach(row => {
            if (row.usn == usn) {
                console.log(row.usn)
                studName = row.name;
                sub5.push(Number(row.subject1));
                sub5.push(Number(row.subject2));
                sub5.push(Number(row.subject3));
                sub5.push(Number(row.subject4));
                sub5.push(Number(row.elective1));
                sub5.push(Number(row.elective2));
                sub5.push(Number(row.lab1));            
                sub5.push(Number(row.lab2));
            }
        })
    }
    else {
        sheet1 = info.worksheets[7];
        const rows1 = await promisify(sheet1.getRows)({
            offset: 6
        })
        rows1.forEach(row => {
            if (row.usn == usn) {
                studName = row.name;
                sub7.push(Number(row.subject1));
                sub7.push(Number(row.subject2));
                sub7.push(Number(row.subject3));
                sub7.push(Number(row.elective1));
                sub7.push(Number(row.elective2));
                sub7.push(Number(row.lab1));            
                sub7.push(Number(row.lab2));
            }
        })
    }
}

//json of marks of student
var s5 = {no:5, sub: sub5};
var s7 = {no:7, sub: sub7};

//to render studentoutput.ejs
app.post('/analysis', (request, response) => {
    usn = request.body.an_usn;
    sem = request.body.an_sem;
    analysisSem(usn,sem);
    if(sem == 5) {      
        response.render('studentoutput', {
            student: s5,
            name: studName
        })
    }
    else {
        response.render('studentoutput', {
            student: s7,
            name: studName
        })
    }
})

//to start it at localhost:3000
app.listen(3000, () => {
    console.log("Server started on port 3000");
})