const mongoose=require("mongoose");
const url="mongodb://localhost:27017/files";
mongoose.connect(url);

mongoose.connect(url, {
    useNewUrlParser: true,
    useUnifiedTopology: true,
    useCreateIndex: true,
});
  
const db = mongoose.connection;

db.on('error', console.error.bind(console, 'MongoDB connection error:'));
db.once('open', () => {
  console.log('Connected to MongoDB!');
});