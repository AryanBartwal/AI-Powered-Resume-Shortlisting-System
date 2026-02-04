// Helper functions for storage operations
const storage = require('./memoryStorage');

module.exports = {
  storage,
  
  // Helper to format job for response
  formatJob: (job) => job,
  
  // Helper to format resume for response
  formatResume: (resume) => resume,
  
  // Helper to format ranking for response
  formatRanking: (ranking) => ranking
};
