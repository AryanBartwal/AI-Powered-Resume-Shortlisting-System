const express = require('express');
const router = express.Router();
const multer = require('multer');
const path = require('path');
const axios = require('axios');
const Resume = require('../models/Resume');
const fs = require('fs');
const { protect, authorize } = require('../middleware/auth');

// Configure multer for file upload
const uploadStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = process.env.UPLOAD_DIR || './uploads';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  }
});

const fileFilter = (req, file, cb) => {
  const allowedTypes = ['application/pdf', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
  if (allowedTypes.includes(file.mimetype)) {
    cb(null, true);
  } else {
    cb(new Error('Only PDF and DOCX files are allowed'), false);
  }
};

const upload = multer({
  storage: uploadStorage,
  fileFilter: fileFilter,
  limits: {
    fileSize: parseInt(process.env.MAX_FILE_SIZE) || 10485760 // 10MB default
  }
});

// Upload resume (Authenticated users)
router.post('/upload', protect, upload.single('resume'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: 'No file uploaded'
      });
    }

    const { candidateName, email, phone, jobId } = req.body;

    if (!candidateName || !email) {
      return res.status(400).json({
        success: false,
        message: 'Candidate name and email are required'
      });
    }

    // Call ML service to parse resume
    const mlServiceUrl = process.env.ML_SERVICE_URL || 'http://localhost:5001';
    let parsedData = {
      text: '',
      skills: [],
      experience: 0,
      education: []
    };

    try {
      const FormData = require('form-data');
      const formData = new FormData();
      formData.append('file', fs.createReadStream(req.file.path), req.file.originalname);

      const parseResponse = await axios.post(`${mlServiceUrl}/api/parse/resume`, formData, {
        headers: formData.getHeaders(),
        timeout: 30000
      });

      if (parseResponse.data.success) {
        parsedData = parseResponse.data.data;
      }
    } catch (mlError) {
      // Continue without parsed data
    }

    // Create resume record
    const resume = await Resume.create({
      candidateName,
      email,
      phone,
      fileName: req.file.originalname,
      filePath: req.file.path,
      fileType: path.extname(req.file.originalname).substring(1),
      parsedContent: parsedData.text,
      skills: parsedData.skills,
      experience: parsedData.experience,
      education: parsedData.education,
      jobId
    });

    res.status(201).json({
      success: true,
      message: 'Resume uploaded successfully',
      data: resume
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error uploading resume',
      error: error.message
    });
  }
});

// Get all resumes for a job
router.get('/job/:jobId', async (req, res) => {
  try {
    const resumes = await Resume.find({ jobId: req.params.jobId }).sort({ createdAt: -1 });
    
    res.json({
      success: true,
      count: resumes.length,
      data: resumes
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error fetching resumes',
      error: error.message
    });
  }
});

// Get resume by ID
router.get('/:id', async (req, res) => {
  try {
    const resume = await Resume.findById(req.params.id);
    
    if (!resume) {
      return res.status(404).json({
        success: false,
        message: 'Resume not found'
      });
    }
    
    res.json({
      success: true,
      data: resume
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error fetching resume',
      error: error.message
    });
  }
});

// Delete resume
router.delete('/:id', async (req, res) => {
  try {
    const resume = await Resume.findById(req.params.id);
    
    if (!resume) {
      return res.status(404).json({
        success: false,
        message: 'Resume not found'
      });
    }

    // Delete file from disk
    if (fs.existsSync(resume.filePath)) {
      fs.unlinkSync(resume.filePath);
    }

    await Resume.findByIdAndDelete(req.params.id);
    
    res.json({
      success: true,
      message: 'Resume deleted successfully'
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error deleting resume',
      error: error.message
    });
  }
});

// Match a resume against all available jobs
router.post('/match-jobs', protect, (req, res, next) => {
  upload.single('resume')(req, res, (err) => {
    if (err) {
      if (err instanceof multer.MulterError) {
        return res.status(400).json({
          success: false,
          message: `File upload error: ${err.message}`
        });
      }
      return res.status(400).json({
        success: false,
        message: err.message || 'Error uploading file'
      });
    }
    next();
  });
}, async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: 'No file uploaded'
      });
    }

    const { candidateName, email, phone } = req.body;

    if (!candidateName || !email) {
      return res.status(400).json({
        success: false,
        message: 'Candidate name and email are required'
      });
    }

    // Call ML service to parse resume
    const mlServiceUrl = process.env.ML_SERVICE_URL || 'http://localhost:5001';
    let parsedData = {
      text: '',
      skills: [],
      experience: 0,
      education: []
    };

    try {
      const FormData = require('form-data');
      const formData = new FormData();
      formData.append('file', fs.createReadStream(req.file.path), req.file.originalname);

      const parseResponse = await axios.post(`${mlServiceUrl}/api/parse/resume`, formData, {
        headers: formData.getHeaders(),
        timeout: 30000
      });

      if (parseResponse.data.success) {
        parsedData = parseResponse.data.data;
      }
    } catch (mlError) {
      // Continue without parsed data
    }

    // Create resume record (without jobId)
    const resume = await Resume.create({
      candidateName,
      email,
      phone,
      fileName: req.file.originalname,
      filePath: req.file.path,
      fileType: path.extname(req.file.originalname).substring(1),
      parsedContent: parsedData.text,
      skills: parsedData.skills,
      experience: parsedData.experience,
      education: parsedData.education
    });

    // Get all available jobs (active or no status field)
    const Job = require('../models/Job');
    const jobs = await Job.find({ 
      $or: [
        { status: 'active' },
        { status: { $exists: false } }
      ]  
    });

    if (jobs.length === 0) {
      return res.json({
        success: true,
        message: 'Resume uploaded successfully',
        data: {
          resume: resume,
          matchedJobs: []
        }
      });
    }

    // Match resume against all jobs
    const matchResults = [];
    
    for (const job of jobs) {
      try {
        const rankingResponse = await axios.post(`${mlServiceUrl}/api/ranking/calculate`, {
          job_description: job.description,
          required_skills: job.requiredSkills || [],
          resumes: [{
            id: resume._id.toString(),
            name: resume.candidateName,
            text: resume.parsedContent || '',
            parsedContent: resume.parsedContent || '',
            skills: resume.skills || [],
            experience: resume.experience || 0
          }]
        }, {
          timeout: 30000
        });

        if (rankingResponse.data.success && rankingResponse.data.data.length > 0) {
          const ranking = rankingResponse.data.data[0];
          matchResults.push({
            jobId: job._id,
            jobTitle: job.title,
            company: job.company,
            location: job.location,
            type: job.type,
            matchPercentage: ranking.match_percentage,
            matchedSkills: ranking.matched_skills,
            missingSkills: ranking.missing_skills,
            similarityScore: ranking.similarity_score
          });
        }
      } catch (matchError) {
        // Continue with other jobs
      }
    }

    // Sort by match percentage (highest first)
    matchResults.sort((a, b) => b.matchPercentage - a.matchPercentage);

    res.status(201).json({
      success: true,
      message: `Resume uploaded and matched against ${matchResults.length} jobs`,
      data: {
        resume: resume,
        matchedJobs: matchResults
      }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: 'Error matching resume with jobs',
      error: error.message
    });
  }
});

module.exports = router;
