import React, { useState, useEffect, useCallback } from 'react';
import axios from 'axios';
import { useAuth } from '../contexts/AuthContext';
import '../styles/Home.css';
import '../styles/JobForm.css';
import '../styles/Dashboard.css';

// --- Icons (Inline SVGs for performance & no-deps) ---
const Icons = {
  Search: (props) => <svg width="20" height="20" fill="none" viewBox="0 0 24 24" stroke="currentColor" {...props}><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>,
  MapPin: (props) => <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" {...props}><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 11a3 3 0 11-6 0 3 3 0 016 0z" /></svg>,
  Briefcase: (props) => <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" {...props}><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 13.255A23.931 23.931 0 0112 15c-3.183 0-6.22-.62-9-1.745M16 6V4a2 2 0 00-2-2h-4a2 2 0 00-2 2v2m4 6h.01M5 20h14a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" /></svg>,
  Upload: (props) => <svg width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="currentColor" {...props}><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" /></svg>,
  Chart: (props) => <svg width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="currentColor" {...props}><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" /></svg>,
  Plus: (props) => <svg width="24" height="24" fill="none" viewBox="0 0 24 24" stroke="currentColor" {...props}><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6v6m0 0v6m0-6h6m-6 0H6" /></svg>,
  ArrowLeft: (props) => <svg width="20" height="20" fill="none" viewBox="0 0 24 24" stroke="currentColor" {...props}><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 19l-7-7m0 0l7-7m-7 7h18" /></svg>,
  Check: (props) => <svg width="20" height="20" fill="none" viewBox="0 0 24 24" stroke="currentColor" {...props}><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" /></svg>
};

function Home() {
  const { user } = useAuth();
  const isAdmin = user?.role === 'admin';
  const [activeTab, setActiveTab] = useState('home');
  const [selectedJob, setSelectedJob] = useState(null);
  const [editingJob, setEditingJob] = useState(null); // For edit mode
  const [applyingToJob, setApplyingToJob] = useState(null); // For job application
  const [jobs, setJobs] = useState([]);
  const [myJobs, setMyJobs] = useState([]); // Store admin's own jobs
  const [jobsLoading, setJobsLoading] = useState(false);
  const [jobsError, setJobsError] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [searchLocation, setSearchLocation] = useState('');
  const [message, setMessage] = useState({ type: '', text: '' });
  const [matchedJobs, setMatchedJobs] = useState([]);

  const fetchJobs = useCallback(async () => {
    setJobsLoading(true);
    setJobsError('');
    try {
      // Add Authorization header if user is logged in
      const config = user ? {
        headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
      } : {};
      
      const response = await axios.get('http://localhost:5000/api/jobs', config);
      if (response.data.success) {
        setJobs(response.data.data);
        
        // If admin, fetch only their own jobs using the createdBy filter
        if (isAdmin && user) {
           const userId = user._id || user.id;
           const myJobsResponse = await axios.get(
             `http://localhost:5000/api/jobs?createdBy=${userId}`, 
             config
           );
           if (myJobsResponse.data.success) {
             setMyJobs(myJobsResponse.data.data);
           }
        }
        return;
      }
      setJobsError(response.data?.message || 'Failed to fetch jobs');
    } catch (error) {
      setJobsError('Failed to fetch jobs. Please try again.');
    } finally {
      setJobsLoading(false);
    }
  }, [user, isAdmin]);

  useEffect(() => {
    fetchJobs();
  }, [fetchJobs]);
  
  const handleDeleteJob = async (jobId, e) => {
    if (e) e.stopPropagation();
    if (!window.confirm('Are you sure you want to delete this job?')) return;

    try {
      const config = {
        headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
      };
      await axios.delete(`http://localhost:5000/api/jobs/${jobId}`, config);
      setMessage({ type: 'success', text: 'Job deleted successfully' });
      fetchJobs();
      if (selectedJob?._id === jobId) setSelectedJob(null);
    } catch (error) {
      setMessage({ type: 'error', text: 'Failed to delete job' });
    }
  };

  const handleEditJob = (job, e) => {
    if (e) e.stopPropagation();
    setEditingJob(job);
    setActiveTab('post-job');
  };

  const normalize = (value) => (value || '').toString().toLowerCase().trim();

  const filteredJobs = jobs.filter((job) => {
    const query = normalize(searchQuery);
    const location = normalize(searchLocation);

    const haystackParts = [
      job?.title,
      job?.company,
      job?.description,
      Array.isArray(job?.requiredSkills) ? job.requiredSkills.join(' ') : job?.requiredSkills,
      job?.type,
    ];
    const haystack = normalize(haystackParts.filter(Boolean).join(' '));
    const jobLocation = normalize(job?.location);

    const matchesQuery = !query || haystack.includes(query);
    const matchesLocation = !location || jobLocation.includes(location);
    return matchesQuery && matchesLocation;
  });

  const handleSearch = () => {
    const jobsEl = document.getElementById('jobs');
    jobsEl?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  const clearFilters = () => {
    setSearchQuery('');
    setSearchLocation('');
  };

  const goHome = () => {
    setActiveTab('home');
    setSelectedJob(null);
    setEditingJob(null);
    setApplyingToJob(null);
    setMessage({ type: '', text: '' });
  };

  return (
    <div className="home-container">
      {activeTab === 'home' && !selectedJob && (
        <>
          {/* Hero Section */}
          <section className="hero-section">
            <h2>Find Your Dream Job Today</h2>
            <p>Connect with top employers and find opportunities that match your skills with AI-powered screening.</p>

            <div className="hero-badges" aria-label="Highlights">
              <span className="hero-badge"><Icons.Check /> Verified Roles</span>
              <span className="hero-badge"><Icons.Check /> AI Match Score</span>
              <span className="hero-badge"><Icons.Check /> Faster Shortlisting</span>
            </div>
            
            <div className="hero-search">
              <input
                type="text"
                className="search-input"
                placeholder="Skills, Designation, Companies"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleSearch();
                }}
                aria-label="Search by skills, designation, or company"
              />
              <input
                type="text"
                className="search-input"
                placeholder="Location"
                value={searchLocation}
                onChange={(e) => setSearchLocation(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleSearch();
                }}
                aria-label="Search by location"
              />
              <button className="search-btn" type="button" onClick={handleSearch}>
                <span className="btn-icon"><Icons.Search /></span>
                Search
              </button>
            </div>
          </section>

          {/* Action Grid */}
          <section className="action-grid">
            {isAdmin && (
              <div className="action-card" onClick={() => { setEditingJob(null); setActiveTab('post-job'); }}>
                <h3><Icons.Plus /> Post a New Job</h3>
                <p>Create a new job listing and let our AI engine find the perfect candidates for you.</p>
              </div>
            )}
            {isAdmin ? (
               <div className="action-card" onClick={() => setActiveTab('my-jobs')}>
                  <h3><Icons.Briefcase /> Manage My Jobs</h3>
                  <p>View, edit, and manage the jobs you have posted.</p>
               </div>
            ) : (
              <div className="action-card" onClick={() => setActiveTab('upload-resume')}>
                <h3><Icons.Upload /> Upload Resume</h3>
                <p>Submit your resume to our database and get matched with multiple job openings instantly.</p>
              </div>
            )}
            {isAdmin && (
              <div className="action-card" onClick={() => setActiveTab('dashboard')}>
                <h3><Icons.Chart /> Ranking Dashboard</h3>
                <p>View detailed insights, candidate rankings, and skill match scores for your posted jobs.</p>
              </div>
            )}
          </section>

          {/* Job Feed */}
          <section className="jobs-section" id="jobs">
            <h3 className="section-title">Latest Opportunities</h3>
            <div className="jobs-toolbar" aria-label="Job results controls">
              <div className="jobs-meta">
                {jobsLoading ? 'Loading jobs…' : `${filteredJobs.length} result${filteredJobs.length === 1 ? '' : 's'}`}
              </div>

              {(searchQuery.trim() || searchLocation.trim()) && (
                <button type="button" className="filter-chip" onClick={clearFilters}>
                  Clear filters
                </button>
              )}
            </div>
            <div className="jobs-grid">
              {jobsLoading ? (
                Array.from({ length: 6 }).map((_, idx) => (
                  <div key={idx} className="job-card skeleton" aria-hidden="true">
                    <div className="job-header">
                      <div className="skeleton-line w-70" />
                      <div className="skeleton-line w-45" />
                    </div>
                    <div className="job-tags">
                      <span className="skeleton-tag" />
                      <span className="skeleton-tag" />
                    </div>
                    <div className="skeleton-line w-100" />
                    <div className="skeleton-line w-95" />
                    <div className="skeleton-line w-80" />
                    <div className="skeleton-line w-60" style={{ marginTop: '20px' }} />
                  </div>
                ))
              ) : jobsError ? (
                <div className="empty-state">
                  <h4>Couldn’t load jobs</h4>
                  <p>{jobsError}</p>
                  <button className="primary-btn" onClick={fetchJobs} type="button" style={{ maxWidth: 220 }}>
                    Retry
                  </button>
                </div>
              ) : filteredJobs.length > 0 ? (
                filteredJobs.map(job => (
                  <div key={job._id} className="job-card clickable-card" onClick={() => setSelectedJob(job)}>
                    <div className="job-header">
                      <h4 className="job-title">{job.title}</h4>
                      <div className="company-name">{job.company || 'Top Tech Company'}</div>
                    </div>
                    <div className="job-tags">
                      {job.location && <span className="tag"><Icons.MapPin /> {job.location}</span>}
                      <span className="tag"><Icons.Briefcase /> {job.type || 'Full-time'}</span>
                      {job.experienceLevel && <span className="tag">{job.experienceLevel}</span>}
                    </div>
                    <p className="job-description line-clamp-3">{job.description}</p>
                    
                    {!isAdmin && (
                      <button 
                        className="apply-btn" 
                        onClick={(e) => {
                          e.stopPropagation();
                          setApplyingToJob(job);
                          setActiveTab('apply-job');
                        }}
                      >
                         Apply Now
                      </button>
                    )}
                  </div>
                ))
              ) : (
                <div className="empty-state">
                  <h4>No matches found</h4>
                  {(searchQuery.trim() || searchLocation.trim()) ? (
                    <p>Try a different keyword or clear the filters.</p>
                  ) : (
                    <p>No active job openings found right now.</p>
                  )}
                  <div className="empty-actions">
                    {(searchQuery.trim() || searchLocation.trim()) && (
                      <button className="primary-btn" onClick={clearFilters} type="button" style={{ maxWidth: 220 }}>
                        Clear filters
                      </button>
                    )}
                    <button className="primary-btn" onClick={() => setActiveTab('post-job')} type="button" style={{ maxWidth: 220 }}>
                      Post a Job
                    </button>
                  </div>
                </div>
              )}
            </div>
          </section>
        </>
      )}

      {/* Job Detail View */}
      {activeTab === 'home' && selectedJob && (
        <JobDetailedView 
          job={selectedJob} 
          isAdmin={isAdmin}
          onBack={() => setSelectedJob(null)}
          onApply={(job) => {
            setSelectedJob(null);
            setActiveTab('upload-resume');
          }}
          onEdit={(job) => {
             setSelectedJob(null);
             handleEditJob(job);
          }}
          onDelete={(jobId) => {
             // Confirm handled in handleDeleteJob logic conceptually, but for now we pass ID
             handleDeleteJob(jobId);
          }}
        />
      )}

      {/* Forms & Dashboard Views */}
      {activeTab !== 'home' && (
        <div style={{ maxWidth: '900px', margin: '0 auto' }}>
          <button onClick={goHome} style={{ 
            marginBottom: '24px', 
            background: 'none', 
            color: 'var(--text-secondary)', 
            fontSize: '1rem', 
            display: 'flex', 
            alignItems: 'center', 
            gap: '8px',
            fontWeight: 600,
            padding: 0
          }}>
            <Icons.ArrowLeft /> Back to Home
          </button>
          
          {activeTab === 'post-job' && (
            <div className="form-container">
              <h2 style={{marginTop: 0, marginBottom: '24px'}}>
                 {editingJob ? 'Edit Job Posting' : 'Post a New Job'}
              </h2>
              <JobPostForm 
                 initialData={editingJob}
                 onJobCreated={() => { 
                    fetchJobs(); 
                    setEditingJob(null);
                    setSelectedJob(null);
                    if (isAdmin) setActiveTab('my-jobs');
                    else setActiveTab('home');
                 }} 
                 setMessage={setMessage} 
              />  
            </div>
          )}

          {activeTab === 'my-jobs' && isAdmin && (
            <div className="form-container my-jobs-container">
               <div className="my-jobs-header">
                 <h2>My Posted Jobs</h2>
               </div>
               
               {myJobs.length === 0 ? (
                 <div className="empty-state">
                   <h4>You haven't posted any jobs yet.</h4>
                   <p>Create your first job posting to start finding candidates.</p>
                 </div>
               ) : (
                 <div className="jobs-grid">
                    {myJobs.map(job => (
                      <div key={job._id} className="job-card static-card">
                        <div className="job-header">
                          <h4 className="job-title">{job.title}</h4>
                          <div className="company-name">{job.company || 'Top Tech Company'}</div>
                        </div>
                        <div className="job-tags">
                           {job.location && <span className="tag"><Icons.MapPin /> {job.location}</span>}
                           <span className="tag"><Icons.Briefcase /> {job.type || 'Full-time'}</span>
                           {job.experienceLevel && <span className="tag">Exp: {job.experienceLevel}</span>}
                           {job.salaryRange && <span className="tag">₹ {job.salaryRange}</span>}
                        </div>
                        <p className="job-description line-clamp-3">{job.description}</p>
                        
                        <div className="admin-actions-footer">
                          <button 
                             className="admin-btn-edit" 
                             onClick={(e) => handleEditJob(job, e)}
                           >
                             Edit
                           </button>
                           <button 
                             className="admin-btn-delete" 
                             onClick={(e) => handleDeleteJob(job._id, e)}
                           >
                             Delete
                           </button>
                        </div>
                      </div>
                    ))}
                 </div>
               )}
            </div>
          )}

          {activeTab === 'upload-resume' && !isAdmin && (
            <div className="form-container" style={{maxWidth: '100%'}}>
              <h2 style={{marginTop: 0, marginBottom: '24px'}}>Upload Your Resume</h2>
              <div style={{display: 'grid', gridTemplateColumns: '1fr 1.5fr', gap: '32px', alignItems: 'start'}}>
                <ResumeUploadForm 
                  jobs={jobs} 
                  setMessage={setMessage} 
                  setMatchedJobs={setMatchedJobs}
                />
                <div>
                  <h3 style={{marginTop: 0, marginBottom: '16px'}}>Jobs Based on Your Resume</h3>
                  {matchedJobs.length > 0 ? (
                    <>
                      <p style={{color: 'var(--text-secondary)', marginBottom: '24px', fontSize: '0.95rem'}}>
                        We found {matchedJobs.length} job{matchedJobs.length !== 1 ? 's' : ''} that match your skills and experience
                      </p>
                      <div style={{display: 'flex', flexDirection: 'column', gap: '16px'}}>
                        {matchedJobs.map((match, index) => (
                          <div key={match.jobId._id} className="ranking-card" style={{cursor: 'pointer'}} onClick={() => setSelectedJob(match.jobId)}>
                            <div className={`rank-badge ${match.matchPercentage >= 75 ? 'high-match' : (match.matchPercentage >= 50 ? 'med-match' : 'low-match')}`}>
                              {match.matchPercentage.toFixed(0)}%
                            </div>
                            <div style={{flex: 1}}>
                              <h4 style={{margin: '0 0 8px 0'}}>{match.jobId.title}</h4>
                              <p style={{margin: '0 0 8px 0', color: 'var(--text-secondary)', fontSize: '0.9rem'}}>
                                {match.jobId.company} • {match.jobId.location}
                              </p>
                              <div style={{fontSize: '0.85rem', color: 'var(--text-tertiary)'}}>
                                <strong>Matched Skills:</strong> {match.matchedSkills?.slice(0, 3).join(', ')}{match.matchedSkills?.length > 3 ? ` +${match.matchedSkills.length - 3}` : ''}
                              </div>
                            </div>
                            <div className="rank-number">#{index + 1}</div>
                          </div>
                        ))}
                      </div>
                    </>
                  ) : (
                    <div style={{
                      textAlign: 'center', 
                      padding: '48px 24px', 
                      color: 'var(--text-tertiary)', 
                      border: '2px dashed var(--border-color)', 
                      borderRadius: '12px',
                      background: 'var(--bg-color)'
                    }}>
                      <Icons.Chart style={{ width: '48px', height: '48px', margin: '0 auto 16px', opacity: 0.5 }} />
                      <h4 style={{margin: '0 0 8px 0', color: 'var(--text-secondary)'}}>No Resume Uploaded Yet</h4>
                      <p style={{margin: 0, fontSize: '0.9rem'}}>
                        Upload your resume to see matching jobs ranked by compatibility
                      </p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {activeTab === 'apply-job' && applyingToJob && !isAdmin && (
            <div className="form-container">
              <h2 style={{marginTop: 0, marginBottom: '8px'}}>Apply for {applyingToJob.title}</h2>
              <p style={{color: 'var(--text-secondary)', marginBottom: '24px'}}>
                {applyingToJob.company} • {applyingToJob.location}
              </p>
              <JobApplicationForm job={applyingToJob} setMessage={setMessage} onSuccess={() => {
                setActiveTab('home');
                setApplyingToJob(null);
              }} />
            </div>
          )}

          {activeTab === 'dashboard' && (
             <div className="form-container" style={{maxWidth: '100%'}}>
               <h2 style={{marginTop: 0, marginBottom: '24px'}}>Candidate Ranking Dashboard</h2>
               <Dashboard jobs={isAdmin ? myJobs : jobs} />
             </div>
          )}
        </div>
      )}

      {message.text && (
        <div className={`message ${message.type}`} style={{ 
          position: 'fixed', 
          bottom: '32px', 
          right: '32px', 
          zIndex: 1100, 
          boxShadow: '0 8px 16px rgba(0,0,0,0.15)',
          minWidth: '300px'
        }}>
          {message.text}
        </div>
      )}
    </div>
  );
}

function JobDetailedView({ job, isAdmin, onBack, onApply, onEdit, onDelete }) {
  if (!job) return null;

  return (
    <div className="job-detail-view">
      {/* Navigation Breadcrumb */}
      <button onClick={onBack} className="back-link">
        <Icons.ArrowLeft /> Back to Jobs
      </button>

      {/* Hero Header for Job */}
      <div className="job-hero">
        <div className="job-main-info">
          <div className="job-title-wrapper">
             <h1>{job.title}</h1>
             <span className="badge-status">{job.type || 'Full-time'}</span>
          </div>
          
          <div className="job-company-row">
            <div className="company-icon-placeholder">
              {(job.company || 'C').charAt(0).toUpperCase()}
            </div>
            <div className="company-details">
              <span className="company-name-large">{job.company || 'Top Tech Company'}</span>
              <span className="location-detail"><Icons.MapPin /> {job.location || 'Remote'}</span>
            </div>
          </div>
        </div>

        <div className="job-actions">
           {!isAdmin && (
             <button className="apply-btn-primary" onClick={() => onApply(job)}>
               Apply Now
             </button>
           )}
        </div>
      </div>

      {/* Quick Stats Grid */}
      <div className="stats-grid">
        <div className="stat-card">
          <span className="stat-label">Experience</span>
          <span className="stat-value">{job.experienceLevel || 'Not specified'}</span> 
        </div>
        <div className="stat-card">
          <span className="stat-label">Salary</span>
          <span className="stat-value">{job.salaryRange || 'Not disclosed'}</span>  
        </div>
        <div className="stat-card">
          <span className="stat-label">Posted On</span>
          <span className="stat-value">{new Date(job.createdAt || Date.now()).toLocaleDateString()}</span>
        </div>
        <div className="stat-card">
          <span className="stat-label">Work Type</span>
          <span className="stat-value">{job.type || 'Full-time'}</span>
        </div>
      </div>

      <div className="job-content-layout">
        <div className="main-content">
          <section className="detail-section">
            <h3 className="section-header">Job Description</h3>
            <p className="description-text">{job.description}</p>
          </section>
        </div>

        <div className="sidebar-content">
          <section className="detail-section sidebar-box">
             <h3 className="section-header small">Required Skills</h3>
             <div className="skills-cloud">
               {job.requiredSkills && job.requiredSkills.map((skill, idx) => (
                 <span key={idx} className="skill-chip">{skill}</span>
               ))}
             </div>
          </section>
        </div>
      </div>
    </div>
  );
}

function JobPostForm({ initialData, onJobCreated, setMessage }) {
  const [formData, setFormData] = useState({
    title: '',
    description: '',
    requiredSkills: '',
    company: '',
    location: '',
    type: 'Full-time',
    experienceLevel: '',
    salaryRange: ''
  });

  useEffect(() => {
    if (initialData) {
      setFormData({
        title: initialData.title || '',
        description: initialData.description || '',
        requiredSkills: Array.isArray(initialData.requiredSkills) 
          ? initialData.requiredSkills.join(', ') 
          : (initialData.requiredSkills || ''),
        company: initialData.company || '',
        location: initialData.location || '',
        type: initialData.type || 'Full-time',
        experienceLevel: initialData.experienceLevel || '',
        salaryRange: initialData.salaryRange || ''
      });
    }
  }, [initialData]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    try {
      const skills = typeof formData.requiredSkills === 'string' 
        ? formData.requiredSkills.split(',').map(s => s.trim()).filter(s => s)
        : formData.requiredSkills;
      
      const config = {
        headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
      };

      const payload = {
        ...formData,
        requiredSkills: skills
      };

      let response;
      if (initialData && initialData._id) {
        // Update existing job
        response = await axios.put(`http://localhost:5000/api/jobs/${initialData._id}`, payload, config);
      } else {
        // Create new job
        response = await axios.post('http://localhost:5000/api/jobs', payload, config);
      }
      
      if (response.data.success) {
        setMessage({ 
          type: 'success', 
          text: initialData ? 'Job updated successfully!' : 'Job posted successfully!' 
        });
        if (!initialData) {
            setFormData({ 
              title: '', description: '', requiredSkills: '', company: '', location: '',
              type: 'Full-time', experienceLevel: '', salaryRange: ''
            });    
        }
        onJobCreated();
      }
    } catch (error) {
      const errorMsg = error.response?.data?.error || error.response?.data?.message || 'Error saving job. Please try again.';
      setMessage({ type: 'error', text: errorMsg });
    }
  };

  return (
    <form onSubmit={handleSubmit} className="job-post-form">
      <div className="form-group">
        <label>Job Title</label>
        <input
          type="text"
          value={formData.title}
          onChange={(e) => setFormData({...formData, title: e.target.value})}
          required
          placeholder="e.g. Senior Software Engineer"
        />
      </div>
      
      <div className="grid-2">
        <div className="form-group">
          <label>Company Name</label>
          <input
            type="text"
            value={formData.company}
            onChange={(e) => setFormData({...formData, company: e.target.value})}
            placeholder="e.g. TechCorp Inc."
          />
        </div>
        <div className="form-group">
          <label>Location</label>
          <input
            type="text"
            value={formData.location}
            onChange={(e) => setFormData({...formData, location: e.target.value})}
            placeholder="e.g. Remote, New York"
          />
        </div>
      </div>

      <div className="grid-3">
        <div className="form-group">
          <label>Job Type</label>
          <select 
            value={formData.type} 
            onChange={(e) => setFormData({...formData, type: e.target.value})}
          >
            <option value="Full-time">Full-time</option>
            <option value="Part-time">Part-time</option>
            <option value="Contract">Contract</option>
            <option value="Internship">Internship</option>
            <option value="Freelance">Freelance</option>
          </select>
        </div>
        <div className="form-group">
          <label>Experience</label>
          <input
            type="text"
            value={formData.experienceLevel}
            onChange={(e) => setFormData({...formData, experienceLevel: e.target.value})}
            placeholder="e.g. 0–2 years"
          />
        </div>
        <div className="form-group">
          <label>Salary Range (per Annum)</label>
          <input
            type="text"
            value={formData.salaryRange}
            onChange={(e) => setFormData({...formData, salaryRange: e.target.value})}
            placeholder="e.g. ₹ 4,50,000 or 10-12 LPA"
          />
        </div>
      </div>

      <div className="form-group">
        <label>Job Description</label>
        <textarea
          value={formData.description}
          onChange={(e) => setFormData({...formData, description: e.target.value})}
          required
          rows="6"
          placeholder="Describe the role, responsibilities, and perks..."
        />
      </div>
      
      <div className="form-group">
        <label>Required Skills (comma-separated)</label>
        <input
          type="text"
          value={formData.requiredSkills}
          onChange={(e) => setFormData({...formData, requiredSkills: e.target.value})}
          placeholder="e.g. React, Node.js, Python, AWS"
        />
      </div>
      
      <button type="submit" className="primary-btn">
        {initialData ? 'Update Job Listing' : 'Create Job Listing'}
      </button>
    </form>
  );
}

// Job Application Form - Apply to a specific job
function JobApplicationForm({ job, setMessage, onSuccess }) {
  const [formData, setFormData] = useState({
    candidateName: '',
    email: '',
    phone: ''
  });
  const [file, setFile] = useState(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!file) {
      setMessage({ type: 'error', text: 'Please select a resume file (PDF or DOCX)' });
      return;
    }

    const uploadData = new FormData();
    uploadData.append('resume', file);
    uploadData.append('jobId', job._id);
    Object.keys(formData).forEach(key => uploadData.append(key, formData[key]));

    try {
      setIsSubmitting(true);
      const config = {
        headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
      };
      const response = await axios.post('http://localhost:5000/api/resumes/upload', uploadData, config);

      if (response.data.success) {
        setMessage({ 
          type: 'success', 
          text: `Application submitted successfully for ${job.title}!` 
        });
        setFormData({ candidateName: '', email: '', phone: '' });
        setFile(null);
        if (onSuccess) onSuccess();
      }
    } catch (error) {
      setMessage({ 
        type: 'error', 
        text: error.response?.data?.message || 'Error submitting application. Please try again.' 
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <form onSubmit={handleSubmit}>
      <div className="form-group">
        <label>Full Name</label>
        <input
          type="text"
          value={formData.candidateName}
          onChange={(e) => setFormData({...formData, candidateName: e.target.value})}
          required
          placeholder="Your full name"
        />
      </div>
      <div className="grid-2">
        <div className="form-group">
          <label>Email Address</label>
          <input
            type="email"
            value={formData.email}
            onChange={(e) => setFormData({...formData, email: e.target.value})}
            required
            placeholder="you@example.com"
          />
        </div>
        <div className="form-group">
          <label>Phone Number</label>
          <input
            type="tel"
            value={formData.phone}
            onChange={(e) => setFormData({...formData, phone: e.target.value})}
            placeholder="+1 (555) 000-0000"
          />
        </div>
      </div>
      <div className="form-group">
        <label>Upload Resume (PDF/DOCX)</label>
        <div className="file-upload-area">
          <input
            type="file"
            accept=".pdf,.docx"
            onChange={(e) => setFile(e.target.files[0])}
            required
          />
          <div style={{ pointerEvents: 'none' }}>
            <Icons.Upload style={{ width: '32px', height: '32px', color: 'var(--primary-color)', marginBottom: '8px' }} />
            <p style={{ margin: 0, fontWeight: 500, color: 'var(--text-secondary)' }}>
              {file ? file.name : 'Click to Upload or Drag & Drop'}
            </p>
            <p style={{marginTop: '4px', fontSize: '0.85rem', color: 'var(--text-tertiary)'}}>
              Maximum file size: 5MB
            </p>
          </div>
        </div>
      </div>
      <button type="submit" className="primary-btn" disabled={isSubmitting}>
        {isSubmitting ? 'Submitting Application...' : 'Submit Application'}
      </button>
    </form>
  );
}

function ResumeUploadForm({ jobs, setMessage, setMatchedJobs }) {
  const [formData, setFormData] = useState({
    candidateName: '',
    email: '',
    phone: ''
  });
  const [file, setFile] = useState(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!file) {
      setMessage({ type: 'error', text: 'Please select a resume file (PDF or DOCX)' });
      return;
    }

    const uploadData = new FormData();
    uploadData.append('resume', file);
    Object.keys(formData).forEach(key => uploadData.append(key, formData[key]));

    try {
      setIsAnalyzing(true);
      setMessage({ type: 'info', text: 'Uploading and analyzing your resume...' });
      
      const config = {
        headers: { 
          'Authorization': `Bearer ${localStorage.getItem('token')}`
        }
      };
      
      const response = await axios.post('http://localhost:5000/api/resumes/match-jobs', uploadData, config);
      
      if (response.data.success) {
        const { matchedJobs } = response.data.data;
        
        // Transform the matched jobs to include full job details
        const enrichedMatches = matchedJobs.map(match => ({
          jobId: {
            _id: match.jobId,
            title: match.jobTitle,
            company: match.company,
            location: match.location,
            type: match.type
          },
          matchPercentage: match.matchPercentage,
          matchedSkills: match.matchedSkills,
          missingSkills: match.missingSkills,
          similarityScore: match.similarityScore
        }));
        
        setMatchedJobs(enrichedMatches);
        
        if (enrichedMatches.length > 0) {
          setMessage({ 
            type: 'success', 
            text: `Resume uploaded successfully! Found ${enrichedMatches.length} matching job${enrichedMatches.length !== 1 ? 's' : ''}.` 
          });
        } else {
          setMessage({ 
            type: 'success', 
            text: 'Resume uploaded successfully! No matching jobs found at the moment.' 
          });
        }
        
        setFormData({ candidateName: '', email: '', phone: '' });
        setFile(null);
      }
    } catch (error) {
      const errorMessage = error.response?.status === 401 
        ? 'Authentication required. Please log in again.'
        : error.response?.data?.message || 'Error uploading resume. Please check file format and try again.';
      
      setMessage({ 
        type: 'error', 
        text: errorMessage
      });
    } finally {
      setIsAnalyzing(false);
    }
  };

  return (
    <form onSubmit={handleSubmit}>
      <div className="form-group">
        <label>Full Name</label>
        <input
          type="text"
          value={formData.candidateName}
          onChange={(e) => setFormData({...formData, candidateName: e.target.value})}
          required
          placeholder="Your full name"
        />
      </div>
      <div className="grid-2">
        <div className="form-group">
          <label>Email Address</label>
          <input
            type="email"
            value={formData.email}
            onChange={(e) => setFormData({...formData, email: e.target.value})}
            required
            placeholder="you@example.com"
          />
        </div>
        <div className="form-group">
          <label>Phone Number</label>
          <input
            type="tel"
            value={formData.phone}
            onChange={(e) => setFormData({...formData, phone: e.target.value})}
            placeholder="+1 (555) 000-0000"
          />
        </div>
      </div>
      <div className="form-group">
        <label>Upload Resume (PDF/DOCX)</label>
        <div className="file-upload-area">
          <input
            type="file"
            accept=".pdf,.docx"
            onChange={(e) => setFile(e.target.files[0])}
            required
          />
          <div style={{ pointerEvents: 'none' }}>
            <Icons.Upload style={{ width: '32px', height: '32px', color: 'var(--primary-color)', marginBottom: '8px' }} />
            <p style={{ margin: 0, fontWeight: 500, color: 'var(--text-secondary)' }}>
              {file ? file.name : 'Click to Upload or Drag & Drop'}
            </p>
            <p style={{marginTop: '4px', fontSize: '0.85rem', color: 'var(--text-tertiary)'}}>
              Maximum file size: 5MB
            </p>
          </div>
        </div>
      </div>
      <button type="submit" className="primary-btn" disabled={isAnalyzing}>
        {isAnalyzing ? 'Analyzing Resume...' : 'Upload & Match Jobs'}
      </button>
    </form>
  );
}

function Dashboard({ jobs }) {
  const [selectedJobId, setSelectedJobId] = useState('');
  const [rankings, setRankings] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const generateRankings = async () => {
    if (!selectedJobId) return;
    setLoading(true);
    setError('');
    try {
      const response = await axios.post(`http://localhost:5000/api/rankings/generate/${selectedJobId}`);
      if (response.data.success) {
        await fetchRankings();
      } else {
        setError(response.data.message || 'Failed to generate rankings');
      }
    } catch (error) {
      const errorMsg = error.response?.data?.message || error.message || 'Failed to generate rankings';
      setError(errorMsg);
    } finally {
      setLoading(false);
    }
  };

  const fetchRankings = useCallback(async () => {
    if (!selectedJobId) return;
    try {
      const response = await axios.get(`http://localhost:5000/api/rankings/job/${selectedJobId}`);     
      if (response.data.success) {
        setRankings(response.data.data);
        setError('');
      }
    } catch (error) {
      setError('Failed to fetch rankings');
    }
  }, [selectedJobId]);

  useEffect(() => {
    if (selectedJobId) {
      fetchRankings();
      return;
    }
    setRankings([]);
    setError('');
  }, [selectedJobId, fetchRankings]);

  return (
    <div className="dashboard">
      <div className="form-group">
        <label>Select Job to Evaluate</label>
        <div className="inline-actions">
          <select 
            value={selectedJobId} 
            onChange={(e) => setSelectedJobId(e.target.value)}
          >
            <option value="">-- Choose a job position --</option>
            {jobs.map(job => (
              <option key={job._id} value={job._id}>{job.title}</option>
            ))}
          </select>
          <button 
            onClick={generateRankings} 
            disabled={!selectedJobId || loading}
            className="primary-btn"
          >
            {loading ? 'Analyzing...' : 'Analyze Candidates'}
          </button>
        </div>
      </div>

      {error && (
        <div className="message error">
           {error}
        </div>
      )}

      {loading && (
        <div style={{textAlign: 'center', padding: '48px', color: 'var(--text-secondary)'}}>
          <div className="spinner" style={{fontSize: '2rem', marginBottom: '16px'}}>⏳</div>
          <p>Analyzing resumes against job requirements...</p>
        </div>
      )}

      {!loading && rankings.length === 0 && selectedJobId && !error && (
        <div style={{
          textAlign: 'center', 
          padding: '48px', 
          color: 'var(--text-tertiary)', 
          border: '2px dashed var(--border-color)', 
          borderRadius: '12px',
          background: 'var(--bg-color)'
        }}>
          <h3>No rankings generated yet</h3>
          <p>Click "Analyze Candidates" to run the AI matching algorithm.</p>
        </div>
      )}

      {rankings.length > 0 && (
        <div className="rankings-list" style={{ marginTop: '32px' }}>
          <div style={{ 
            display: 'flex', 
            justifyContent: 'space-between', 
            alignItems: 'center', 
            marginBottom: '16px',
            borderBottom: '1px solid var(--border-color)',
            paddingBottom: '16px'
          }}>
            <h3 style={{margin: 0}}>Top Candidates ({rankings.length})</h3>
            <span style={{fontSize: '0.9rem', color: 'var(--text-secondary)'}}>Sorted by Match Score</span>
          </div>
          
          {rankings.map((ranking, index) => {
             const score = ranking.matchPercentage || 0;
             const matchClass = score >= 75 ? 'high-match' : (score >= 50 ? 'med-match' : 'low-match');
             const rankLabel = index === 0 ? '🥇' : index === 1 ? '🥈' : index === 2 ? '🥉' : `#${index + 1}`;
             
             return (
              <div key={ranking._id} className={`ranking-card rank-${index + 1}`}>
                <div className="rank-indicator">
                  <div className="rank-medal">{rankLabel}</div>
                  <div className="rank-text">Rank {index + 1}</div>
                </div>
                
                <div className={`score-badge ${matchClass}`}>
                  <div className="score-value">{score.toFixed(0)}%</div>
                  <div className="score-label">Match</div>
                </div>
                
                <div className="candidate-details">
                  <h4 className="candidate-name">{ranking.resumeId?.candidateName || 'Unknown Candidate'}</h4>
                  <p className="candidate-email">{ranking.resumeId?.email}</p>
                  {ranking.resumeId?.phone && (
                    <p className="candidate-phone">{ranking.resumeId?.phone}</p>
                  )}
                </div>
                
                <div className="skills-info">
                  <div className="skills-header">
                    <span className="skills-icon">✓</span>
                    <span>Matched Skills</span>
                  </div>
                  <div className="skills-list">
                    {ranking.matchedSkills?.length > 0 ? (
                      <>
                        {ranking.matchedSkills.slice(0, 4).map((skill, idx) => (
                          <span key={idx} className="skill-tag">{skill}</span>
                        ))}
                        {ranking.matchedSkills.length > 4 && (
                          <span className="skill-tag more">+{ranking.matchedSkills.length - 4} more</span>
                        )}
                      </>
                    ) : (
                      <span className="no-skills">No specific matches</span>
                    )}
                  </div>
                </div>
              </div>
             );
          })}
        </div>
      )}
    </div>
  );
}

export default Home;
