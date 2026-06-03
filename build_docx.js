/**
 * Generate ATS-friendly Word resume for Rayudu Sai Charan Teja.
 * Run: npm install && node build_docx.js
 */

const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  HeadingLevel,
  LevelFormat,
} = require("docx");

const NAVY = "1A365D";
const BLUE = "2C5282";

function heading(text) {
  return new Paragraph({
    spacing: { before: 240, after: 80 },
    children: [
      new TextRun({ text: text.toUpperCase(), bold: true, size: 22, color: NAVY }),
    ],
  });
}

function jobTitle(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 20 })],
  });
}

function jobCompany(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 20, color: BLUE })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 40 },
    children: [new TextRun({ text, size: 20 })],
  });
}

function body(text) {
  return new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 20 })],
  });
}

function skillLine(category, items) {
  return new Paragraph({
    spacing: { after: 40 },
    children: [
      new TextRun({ text: `${category}: `, bold: true, size: 20 }),
      new TextRun({ text: items, size: 20 }),
    ],
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 360, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          margin: { top: 720, bottom: 720, left: 864, right: 864 },
        },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "Rayudu Sai Charan Teja",
              bold: true,
              size: 36,
              color: NAVY,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [
            new TextRun({
              text: "Performance Marketing Manager | Digital Marketing Strategist | AI Marketing & Technology Consultant",
              bold: true,
              size: 22,
              color: BLUE,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 40 },
          children: [
            new TextRun({
              text: "Hyderabad, India | +91 9000148428 | +91 9390438428 | Saicharanteja47@gmail.com | linkedin.com/in/rayudu-sai-charan-teja-3465a12bb | github.com/rayuduchara-2005",
              size: 18,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 160 },
          children: [
            new TextRun({
              text: "Open to Relocation: Dubai, Abu Dhabi, Saudi Arabia, Qatar | Available for UAE & GCC Opportunities",
              size: 18,
              italics: true,
            }),
          ],
        }),

        heading("Professional Summary"),
        body(
          "Performance Marketing Manager and Technology Consultant with 4+ years driving measurable growth through paid advertising, SEO, lead generation, and AI-powered marketing automation. Proven record managing high-budget Google, Meta, and YouTube ad campaigns, building conversion-focused websites, and delivering data-driven results that lower cost-per-lead and increase qualified pipeline. Combines marketing strategy with hands-on web development and emerging AI tools to maximize ROI for startups and enterprises across India and international markets."
        ),

        heading("Professional Experience"),
        jobTitle("Digital Marketing & Technology Consultant"),
        jobCompany("SMR Oneway  Software Development & IT Infrastructure Solutions | 2021  Present"),
        ...[
          "Managed end-to-end Google Ads and Meta Ads campaigns, achieving 3545% reduction in cost-per-lead while increasing qualified lead volume by 60%+ across client portfolios",
          "Planned and executed comprehensive SEO strategies including technical, on-page, and off-page optimization, improving organic search visibility and keyword rankings",
          "Generated consistent business leads through multi-channel digital marketing including paid search, social advertising, email campaigns, and conversion-optimized landing pages",
          "Coordinated website and mobile application development projects from requirements gathering through deployment, ensuring alignment with marketing and business growth objectives",
          "Managed client communications, project delivery timelines, and stakeholder reporting with a focus on transparency and measurable outcomes",
          "Improved online brand visibility and conversion rates through A/B testing, landing page optimization, and data-driven campaign adjustments using Google Analytics and Tag Manager",
          "Implemented marketing automation workflows and AI-powered content strategies to scale campaign production and improve engagement metrics",
        ].map(bullet),

        heading("Key Projects"),
        jobTitle("WaveXOne Trading Platform (UAE)"),
        jobCompany("Digital Marketing & Development Lead | wavexone.com"),
        ...[
          "Developed and optimized a complete trading platform website with modern UI/UX, responsive design, and conversion-focused user journeys",
          "Implemented SEO optimization strategy improving search visibility for financial services and trading-related keywords in UAE market",
          "Integrated lead generation funnels and conversion-focused marketing strategies to capture and nurture qualified trading platform leads",
          "Managed deployment, Cloudflare CDN setup, security enhancements, SSL configuration, and platform performance optimization",
          "Coordinated cross-functional delivery of marketing campaigns aligned with platform launch and user acquisition goals",
        ].map(bullet),

        jobTitle("DealDunia MLM Platform"),
        jobCompany("Project Management, Marketing & Technical Coordination"),
        ...[
          "Contributed to MLM web application and Android mobile application development with referral-based growth architecture",
          "Managed marketing strategy and business growth initiatives including user acquisition campaigns and referral program optimization",
          "Coordinated API integration and platform functionality ensuring seamless user experience across web and mobile channels",
          "Assisted in lead generation and user acquisition campaigns, driving scalable membership growth through digital channels",
          "Worked with scalable business and referral-based growth systems to support multi-tier network marketing operations",
        ].map(bullet),

        heading("Core Competencies"),
        skillLine(
          "Performance Marketing",
          "Google Ads, Meta Ads, YouTube Ads, SEM, PPC, Lead-Gen Campaigns, Funnel Design, Remarketing, Budget Optimization, Audience Segmentation, Conversion Rate Optimization"
        ),
        skillLine(
          "SEO",
          "Technical, On-Page, Off-Page, Local & International SEO, Keyword Research, Link Building, Schema, Core Web Vitals, SEO Audits, Google Search Console"
        ),
        skillLine(
          "Analytics & Data",
          "GA4, Google Tag Manager, Looker Studio, Conversion & Event Tracking, UTM Management, Attribution, KPI Reporting"
        ),
        skillLine(
          "Web & Development",
          "HTML5, CSS3, JavaScript, PHP, WordPress, Responsive Design, Landing Pages, API Integration, Cloudflare, Hosting & Deployment"
        ),
        skillLine(
          "AI & Automation",
          "ChatGPT, Claude, Cursor, Gemini, Perplexity, Copilot, Jasper, Midjourney, Synthesia, Prompt Engineering, Workflow Automation"
        ),
        skillLine(
          "Design & Social",
          "Canva, Adobe Photoshop, Premiere Pro, Figma, UI/UX Fundamentals, Facebook, Instagram, LinkedIn & YouTube Strategy"
        ),
        skillLine(
          "Tools",
          "GitHub, VS Code, Cursor, Google Workspace, MS Office, Notion, Trello, Asana, Slack"
        ),

        heading("Certifications"),
        ...[
          "Digital Marketing, SEO & Google Ads Certification",
          "Front-End Web Development (HTML, CSS, JavaScript)",
          "UI/UX Basics & Responsive Web Design",
          "AI-Powered Marketing & Content Automation",
        ].map(bullet),

        heading("Education"),
        jobTitle("Bachelor's Degree"),
        body("Andhra University College, Ramachandrapuram"),
        jobTitle("Intermediate Education"),
        body("Board of Intermediate Education, Ramachandrapuram, AP"),
        jobTitle("Secondary School (10th)"),
        body("ZPHS School, Andhra Pradesh"),

        heading("Languages"),
        ...[
          "English  Professional",
          "Hindi  Professional",
          "Telugu  Native",
        ].map(bullet),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  const out = "Rayudu_Sai_Charan_Teja_Resume.docx";
  fs.writeFileSync(out, buffer);
  console.log(`Resume saved to ${out}`);
});
