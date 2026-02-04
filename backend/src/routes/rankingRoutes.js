const express = require('express');
const router = express.Router();
const axios = require('axios');
const Ranking = require('../models/Ranking');
const Job = require('../models/Job');
const Resume = require('../models/Resume');
const { protect, authorize } = require('../middleware/auth');

// Generate rankings for a job (Admin only)
router.post('/generate/:jobId', protect, authorize('admin'), async (req, res) => {
  try {
    const { jobId } = req.params;

    // Get job details
    const job = await Job.findById(jobId);
    if (!job) {
      return res.status(404).json({
        success: false,
        message: 'Job not found'
      });
    }

    // Get all resumes for this job
    const resumes = await Resume.find({ jobId });
    if (resumes.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'No resumes found for this job'
      });
    }

    // Call ML service to calculate similarity scores
    const mlServiceUrl = process.env.ML_SERVICE_URL || 'http://localhost:5001';
    
    try {
      const rankingResponse = await axios.post(`${mlServiceUrl}/api/ranking/calculate`, {
        job_description: job.description,
        required_skills: job.requiredSkills || [],
        resumes: resumes.map(r => ({
          id: r._id.toString(),
          name: r.candidateName,
          text: r.parsedContent || '',
          parsedContent: r.parsedContent || '',
          skills: r.skills || [],
          experience: r.experience || 0
        }))
      }, {
        timeout: 60000
      });

      if (!rankingResponse.data.success) {
        throw new Error('ML service failed to calculate rankings');
      }

      const rankings = rankingResponse.data.data;

      // Delete existing rankings for this job
      await Ranking.deleteMany({ jobId });

      // Create new ranking records
      const rankingRecords = rankings.map((ranking, index) => ({
        jobId,
        resumeId: ranking.resume_id,
        similarityScore: ranking.similarity_score,
        matchPercentage: ranking.match_percentage,
        matchedSkills: ranking.matched_skills,
        missingSkills: ranking.missing_skills,
        rank: index + 1
      }));

      const savedRankings = await Ranking.insertMany(rankingRecords);

      res.json({
        success: true,
        message: 'Rankings generated successfully',
        data: savedRankings
      });
    } catch (mlError) {
      return res.status(500).json({
        success: false,
        message: 'Error generating rankings from ML service',
        error: mlError.message
      });
    }
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error generating rankings',
      error: error.message
    });
  }
});

// Get rankings for a job
router.get('/job/:jobId', async (req, res) => {
  try {
    const { jobId } = req.params;
    const { minScore } = req.query;

    const filter = { jobId };
    if (minScore) {
      filter.matchPercentage = { $gte: parseFloat(minScore) };
    }

    const rankings = await Ranking.find(filter)
      .populate('resumeId', 'candidateName email skills experience')
      .sort({ rank: 1 });

    res.json({
      success: true,
      count: rankings.length,
      data: rankings
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error fetching rankings',
      error: error.message
    });
  }
});

// Get top N ranked resumes
router.get('/job/:jobId/top/:n', async (req, res) => {
  try {
    const { jobId, n } = req.params;

    const rankings = await Ranking.find({ jobId })
      .populate('resumeId', 'candidateName email skills experience')
      .sort({ rank: 1 })
      .limit(parseInt(n));

    res.json({
      success: true,
      count: rankings.length,
      data: rankings
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error fetching top rankings',
      error: error.message
    });
  }
});

// Update ranking status
router.patch('/:id/status', async (req, res) => {
  try {
    const { status } = req.body;

    if (!['pending', 'shortlisted', 'rejected'].includes(status)) {
      return res.status(400).json({
        success: false,
        message: 'Invalid status'
      });
    }

    const ranking = await Ranking.findByIdAndUpdate(
      req.params.id,
      { status },
      { new: true }
    ).populate('resumeId', 'candidateName email');

    if (!ranking) {
      return res.status(404).json({
        success: false,
        message: 'Ranking not found'
      });
    }

    res.json({
      success: true,
      message: 'Ranking status updated',
      data: ranking
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error updating ranking status',
      error: error.message
    });
  }
});

module.exports = router;
