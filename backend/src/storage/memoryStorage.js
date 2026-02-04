// In-memory storage for jobs, resumes, and rankings (no database required)
class MemoryStorage {
  constructor() {
    this.jobs = new Map();
    this.resumes = new Map();
    this.rankings = new Map();
    this.jobCounter = 1;
    this.resumeCounter = 1;
    this.rankingCounter = 1;
  }

  // Job operations
  createJob(jobData) {
    const id = String(this.jobCounter++);
    const job = {
      _id: id,
      ...jobData,
      createdAt: new Date(),
      updatedAt: new Date(),
      status: jobData.status || 'active'
    };
    this.jobs.set(id, job);
    return job;
  }

  getAllJobs(filter = {}) {
    let jobs = Array.from(this.jobs.values());
    if (filter.status) {
      jobs = jobs.filter(job => job.status === filter.status);
    }
    return jobs.sort((a, b) => b.createdAt - a.createdAt);
  }

  getJobById(id) {
    return this.jobs.get(id) || null;
  }

  updateJob(id, updates) {
    const job = this.jobs.get(id);
    if (!job) return null;
    const updatedJob = { ...job, ...updates, updatedAt: new Date() };
    this.jobs.set(id, updatedJob);
    return updatedJob;
  }

  deleteJob(id) {
    return this.jobs.delete(id);
  }

  // Resume operations
  createResume(resumeData) {
    const id = String(this.resumeCounter++);
    const resume = {
      _id: id,
      ...resumeData,
      uploadedAt: new Date()
    };
    this.resumes.set(id, resume);
    return resume;
  }

  getResumesByJobId(jobId) {
    return Array.from(this.resumes.values())
      .filter(resume => resume.jobId === jobId)
      .sort((a, b) => b.uploadedAt - a.uploadedAt);
  }

  getResumeById(id) {
    const resume = this.resumes.get(id);
    if (resume && resume.jobId) {
      const job = this.jobs.get(resume.jobId);
      return { ...resume, jobId: job };
    }
    return resume || null;
  }

  deleteResume(id) {
    return this.resumes.delete(id);
  }

  // Ranking operations
  createRanking(rankingData) {
    const id = String(this.rankingCounter++);
    const ranking = {
      _id: id,
      ...rankingData,
      createdAt: new Date(),
      status: rankingData.status || 'pending'
    };
    this.rankings.set(id, ranking);
    return ranking;
  }

  createManyRankings(rankingsData) {
    return rankingsData.map(data => this.createRanking(data));
  }

  getRankingsByJobId(jobId, filter = {}) {
    let rankings = Array.from(this.rankings.values())
      .filter(ranking => ranking.jobId === jobId);

    if (filter.minScore) {
      rankings = rankings.filter(r => r.matchPercentage >= filter.minScore);
    }

    // Populate resume and job data
    rankings = rankings.map(ranking => {
      const resume = this.resumes.get(ranking.resumeId);
      const job = this.jobs.get(ranking.jobId);
      return {
        ...ranking,
        resumeId: resume,
        jobId: job
      };
    });

    return rankings.sort((a, b) => a.rank - b.rank);
  }

  updateRanking(id, updates) {
    const ranking = this.rankings.get(id);
    if (!ranking) return null;
    const updatedRanking = { ...ranking, ...updates };
    this.rankings.set(id, updatedRanking);
    return updatedRanking;
  }

  deleteRankingsByJobId(jobId) {
    const toDelete = [];
    for (const [id, ranking] of this.rankings.entries()) {
      if (ranking.jobId === jobId) {
        toDelete.push(id);
      }
    }
    toDelete.forEach(id => this.rankings.delete(id));
    return toDelete.length;
  }

  // Utility methods
  clearAll() {
    this.jobs.clear();
    this.resumes.clear();
    this.rankings.clear();
    this.jobCounter = 1;
    this.resumeCounter = 1;
    this.rankingCounter = 1;
  }
}

// Singleton instance
const storage = new MemoryStorage();

module.exports = storage;
