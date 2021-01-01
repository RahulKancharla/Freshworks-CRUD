
var fs =require('fs')
//create file
var data=fs.readFileSync('student.json')
var students=JSON.parse(data)
console.log(students)
var express=require('express')
var app=express()
var port=9000;
app.listen({port},()=>{
    console.log('server started')
});

app.use(express.static('website'));
//add data to file
app.get('/add/:student/:marks',addMarks);
function addMarks(req,res){
    var data=req.params;
    var student=data.student;
    var marks=Number(data.marks);
    students[student]=marks;
    var info=JSON.stringify(students,null,2)
    fs.writeFile('student.json',info,finished)
    function finished(err){
        console.log(err);
        reply={
            student:student,
            mark:marks,
            status:"success"
        }
        res.send(reply)
    }
    //res.send('student marks  added')
}
//search data 
app.get('/search/:student',searchstudent)
function searchstudent(req,res){
    var student=req.params.student;
    var reply;
    if(students[student]){
         reply={
             status:"found",
             student:student,
            marks:students[student]
             }

    }else{
        reply={
            status:"not found",
            student:student

    }
    
}
res.send(reply)
}
//read all data
app.get('/all',sendall);
function sendall(req,res){
    res.send(students);
}
//delete elements in json file
app.get('/delete/:student',deleteStudent);
function deleteStudent(req,res){
    var student=req.params.student;
    delete students[student];
    console.log(students)
}
