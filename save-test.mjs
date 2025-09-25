import { MongoClient } from 'mongodb';

const uri = 'mongodb+srv://paint:2546paint@cluster0.nmil3pa.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0';

async function saveTestData() {
  const client = new MongoClient(uri);
  
  try {
    console.log('Connecting to MongoDB Atlas...');
    await client.connect();
    console.log('✅ Connected!');
    
    const db = client.db('feedback');
    const collection = db.collection('submissions');
    
    const testData = {
      name: 'John Test',
      email: 'john@test.com',
      rating: '5',
      feedback: 'This is a test from local machine',
      timestamp: new Date()
    };
    
    console.log('Inserting test data...');
    const result = await collection.insertOne(testData);
    console.log('✅ Data saved with ID:', result.insertedId);
    
    // Verify by reading it back
    const saved = await collection.findOne({ _id: result.insertedId });
    console.log('✅ Verified data:', saved);
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  } finally {
    await client.close();
    console.log('Connection closed');
  }
}

saveTestData();
