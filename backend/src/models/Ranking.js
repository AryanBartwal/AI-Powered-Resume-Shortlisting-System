const mongoose = require('mongoose');

const rankingSchema = new mongoose.Schema({
  jobId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'Job',
    required: true
  },
  resumeId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'Resume',
    required: true
  },
  rank: {
    type: Number,
    required: true
  },
  matchPercentage: {
    type: Number,
    required: true,
    min: 0,
    max: 100
  },
  similarityScore: {
    type: Number,
    default: 0
  },
  matchedSkills: [{
    type: String,
    trim: true
  }],
  missingSkills: [{
    type: String,
    trim: true
  }],
  notes: {
    type: String
  }
}, {
  timestamps: true
});

// Compound index for unique ranking per job-resume pair
rankingSchema.index({ jobId: 1, resumeId: 1 }, { unique: true });
rankingSchema.index({ jobId: 1, rank: 1 });

module.exports = mongoose.model('Ranking', rankingSchema);
