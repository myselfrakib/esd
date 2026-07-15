// Presentation State
let currentSlide = 0;
const slides = document.querySelectorAll('.slide');
const slideIndicator = document.getElementById('slide-indicator');
const prevBtn = document.getElementById('prev-btn');
const nextBtn = document.getElementById('next-btn');
const progressBar = document.getElementById('progress-bar');
const fullscreenBtn = document.getElementById('fullscreen-btn');
const downloadBtn = document.getElementById('download-btn');

const totalSlides = slides.length;

// Initialize
function initPresentation() {
  showSlide(currentSlide);
  
  // Navigation Event Listeners
  prevBtn.addEventListener('click', navigatePrev);
  nextBtn.addEventListener('click', navigateNext);
  
  // Keyboard Navigation
  document.addEventListener('keydown', handleKeyDown);
  
  // Fullscreen Button
  fullscreenBtn.addEventListener('click', toggleFullscreen);
  
  // Download PPTX Button
  downloadBtn.addEventListener('click', generatePowerPoint);
}

// Show Specific Slide
function showSlide(index) {
  // Bound check
  if (index < 0) index = 0;
  if (index >= totalSlides) index = totalSlides - 1;
  
  // Update state
  currentSlide = index;
  
  // Update visibility classes
  slides.forEach((slide, idx) => {
    if (idx === currentSlide) {
      slide.classList.add('active');
    } else {
      slide.classList.remove('active');
    }
  });
  
  // Update Indicator & Progress Bar
  slideIndicator.textContent = `${currentSlide + 1} / ${totalSlides}`;
  progressBar.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
  
  // Enable/Disable buttons
  prevBtn.disabled = currentSlide === 0;
  nextBtn.disabled = currentSlide === totalSlides - 1;
}

// Navigation Functions
function navigateNext() {
  if (currentSlide < totalSlides - 1) {
    showSlide(currentSlide + 1);
  }
}

function navigatePrev() {
  if (currentSlide > 0) {
    showSlide(currentSlide - 1);
  }
}

// Keyboards Handler
function handleKeyDown(e) {
  switch (e.key) {
    case 'ArrowRight':
    case ' ':
    case 'Enter':
      e.preventDefault();
      navigateNext();
      break;
    case 'ArrowLeft':
    case 'Backspace':
      e.preventDefault();
      navigatePrev();
      break;
    case 'Home':
      e.preventDefault();
      showSlide(0);
      break;
    case 'End':
      e.preventDefault();
      showSlide(totalSlides - 1);
      break;
  }
}

// Fullscreen Control
function toggleFullscreen() {
  const container = document.documentElement;
  
  if (!document.fullscreenElement) {
    container.requestFullscreen().then(() => {
      fullscreenBtn.querySelector('span').textContent = 'Exit';
      fullscreenBtn.querySelector('i').className = 'fa-solid fa-compress';
    }).catch(err => {
      console.error(`Error attempting to enable fullscreen: ${err.message}`);
    });
  } else {
    document.exitFullscreen().then(() => {
      fullscreenBtn.querySelector('span').textContent = 'Fullscreen';
      fullscreenBtn.querySelector('i').className = 'fa-solid fa-expand';
    });
  }
}

// Fullscreen Change Event
document.addEventListener('fullscreenchange', () => {
  if (!document.fullscreenElement) {
    fullscreenBtn.querySelector('span').textContent = 'Fullscreen';
    fullscreenBtn.querySelector('i').className = 'fa-solid fa-expand';
  }
});

// PowerPoint (.pptx) Export Functionality using PptxGenJS
function generatePowerPoint() {
  // Show loading indicator on button
  const originalText = downloadBtn.querySelector('span').innerHTML;
  downloadBtn.querySelector('span').textContent = 'Generating...';
  downloadBtn.disabled = true;

  try {
    // 1. Initialize PptxGenJS presentation
    const pptx = new PptxGenJS();
    
    // Set presentation properties
    pptx.title = "Sustainable Cities and Communities (SDG 11)";
    pptx.subject = "UN Sustainable Development Goal 11";
    pptx.author = "EcoUrban Presenter";
    pptx.layout = 'LAYOUT_16x9';

    // Theme Palette Definitions
    const BG_COLOR = "0B0F19";      // Slate 950 (Dark theme base)
    const CARD_BG = "151E2E";       // Slate 900
    const ACCENT_BG = "1A1C38";     // Purple/Indigo Callout
    const ECO_BG = "122C20";        // Dark Forest Green
    const ECO_BORDER = "1E5E3D";    // Green Border
    
    const COLOR_PRIMARY = "10B981";  // Emerald Green
    const COLOR_SECONDARY = "14B8A6";// Teal
    const COLOR_ACCENT = "6366F1";   // Indigo
    const COLOR_TEXT = "F8FAFC";     // Off-white
    const COLOR_MUTED = "94A3B8";    // Slate Gray

    // Font definitions
    const FONT_TITLE = "Outfit";
    const FONT_BODY = "Arial";

    // Reusable Header Helper for Content Slides
    function addSlideHeader(slide, titleText, subtitleText, badgeText = "SDG 11: SUSTAINABLE CITIES") {
      // Add Left Emerald Accent Ribbon
      slide.addShape('rect', { x: 0, y: 0, w: 0.12, h: 5.625, fill: COLOR_PRIMARY });

      // Add Badge
      slide.addText(badgeText, {
        x: 0.5, y: 0.3, w: 9.0, h: 0.25,
        fontSize: 9, bold: true, color: COLOR_PRIMARY,
        fontFace: FONT_TITLE, characterSpacing: 1
      });

      // Add Main Slide Title
      slide.addText(titleText, {
        x: 0.5, y: 0.5, w: 9.0, h: 0.45,
        fontSize: 20, bold: true, color: COLOR_TEXT,
        fontFace: FONT_TITLE
      });

      // Add Subtitle / Lead text if provided
      if (subtitleText) {
        slide.addText(subtitleText, {
          x: 0.5, y: 0.95, w: 9.0, h: 0.35,
          fontSize: 10, color: COLOR_MUTED,
          fontFace: FONT_BODY
        });
      }
    }

    // ------------------------------------------------------------------------
    // SLIDE 1: Title Slide
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };

      // Left Accent Ribbon
      slide.addShape('rect', { x: 0, y: 0, w: 0.15, h: 5.625, fill: COLOR_PRIMARY });

      // Badge
      slide.addText("SDG 11: SUSTAINABLE DEVELOPMENT GOAL", {
        x: 0.6, y: 1.0, w: 4.5, h: 0.3,
        fontSize: 9, bold: true, color: COLOR_PRIMARY,
        fontFace: FONT_TITLE, characterSpacing: 1
      });

      // Title
      slide.addText("Sustainable Cities &\nCommunities", {
        x: 0.6, y: 1.4, w: 4.5, h: 1.4,
        fontSize: 32, bold: true, color: COLOR_TEXT,
        fontFace: FONT_TITLE
      });

      // Subtitle
      slide.addText("Shaping the Urban Future through Innovation, Ecology, and Equity", {
        x: 0.6, y: 2.9, w: 4.5, h: 0.7,
        fontSize: 13, color: COLOR_MUTED,
        fontFace: FONT_BODY
      });

      // Horizontal Divider Line
      slide.addShape('rect', { x: 0.6, y: 3.7, w: 2.0, h: 0.03, fill: COLOR_PRIMARY });

      // Footer Meta
      slide.addText("July 2026  |  Global Agenda", {
        x: 0.6, y: 3.9, w: 4.5, h: 0.3,
        fontSize: 9.5, color: COLOR_ACCENT,
        fontFace: FONT_BODY
      });

      // Right image
      slide.addImage({
        path: "images/slide1_hero.png",
        x: 5.3, y: 1.0, w: 4.2, h: 3.6
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 2: The Urban Surge
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "The Urban Surge", "Understanding the massive scale of global urbanization.");

      // Left Column Content
      slide.addText("Today, more than half of the world's population lives in cities. By 2050, that figure is projected to rise to nearly 70%.", {
        x: 0.5, y: 1.5, w: 4.5, h: 0.9,
        fontSize: 12, bold: true, color: COLOR_TEXT,
        fontFace: FONT_BODY, lineSpacing: 16
      });

      slide.addText("Cities are the engines of economic growth, accounting for about 80% of global GDP. However, they also consume over two-thirds of the world's energy and account for more than 70% of global carbon emissions.", {
        x: 0.5, y: 2.5, w: 4.5, h: 1.2,
        fontSize: 10, color: COLOR_MUTED,
        fontFace: FONT_BODY, lineSpacing: 14
      });

      // Stats row inside left column
      const stats = [
        { num: "4.4B", label: "In cities today" },
        { num: "70%", label: "Of global CO2" },
        { num: "60%", label: "Urban by 2030" }
      ];

      stats.forEach((stat, idx) => {
        const xPos = 0.5 + idx * 1.55;
        // Background card shape
        slide.addShape('rect', {
          x: xPos, y: 3.8, w: 1.4, h: 1.1,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Stat number
        slide.addText(stat.num, {
          x: xPos + 0.05, y: 3.9, w: 1.3, h: 0.35,
          fontSize: 16, bold: true, color: COLOR_PRIMARY,
          fontFace: FONT_TITLE, align: 'center'
        });
        // Stat label
        slide.addText(stat.label.toUpperCase(), {
          x: xPos + 0.05, y: 4.3, w: 1.3, h: 0.5,
          fontSize: 7.5, bold: true, color: COLOR_MUTED,
          fontFace: FONT_BODY, align: 'center'
        });
      });

      // Right Column - Skyline Image
      slide.addImage({
        path: "images/slide2_urbanization.png",
        x: 5.3, y: 1.5, w: 4.2, h: 3.4
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 3: Core Pillars of SDG 11
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Core Pillars of Sustainable Urbanism", "Achieving SDG 11 requires a holistic approach addressing multiple interconnected systems.");

      // Left Column content & Image
      slide.addImage({
        path: "images/slide3_pillars.png",
        x: 0.5, y: 1.6, w: 4.2, h: 3.3
      });

      // Right Column 2x2 pillars grid
      const pillars = [
        { title: "Housing & Services", desc: "Ensuring access to safe, affordable, and basic housing globally." },
        { title: "Sustainable Mobility", desc: "Expanding safe, affordable, and accessible public transit networks." },
        { title: "Urban Resilience", desc: "Improving disaster preparedness and mitigating environmental impacts." },
        { title: "Circular Flow", desc: "Reducing municipal waste and supporting localized green economies." }
      ];

      pillars.forEach((pillar, idx) => {
        // Calculate 2x2 positions
        const row = Math.floor(idx / 2);
        const col = idx % 2;
        const xPos = 5.3 + col * 2.15;
        const yPos = 1.6 + row * 1.65;

        // Card Background
        slide.addShape('rect', {
          x: xPos, y: yPos, w: 2.05, h: 1.5,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Title
        slide.addText(pillar.title, {
          x: xPos + 0.1, y: yPos + 0.1, w: 1.85, h: 0.35,
          fontSize: 11, bold: true, color: COLOR_SECONDARY,
          fontFace: FONT_TITLE
        });
        // Description
        slide.addText(pillar.desc, {
          x: xPos + 0.1, y: yPos + 0.45, w: 1.85, h: 0.95,
          fontSize: 8.5, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 12
        });
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 4: Rethinking Transit & Urban Mobility
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Rethinking Transit & Urban Mobility", "Transitioning away from private vehicle dependency toward shared, electric, and active transportation.");

      // Left Column - Bullet list
      slide.addText([
        { text: "Integrated Mass Transit\n", options: { bold: true, color: COLOR_TEXT, fontSize: 11 } },
        { text: "High-capacity rapid transit networks powered by renewable electricity.\n\n", options: { color: COLOR_MUTED, fontSize: 9.5 } },
        
        { text: "Active Infrastructure\n", options: { bold: true, color: COLOR_TEXT, fontSize: 11 } },
        { text: "Protected cycling networks and pedestrianized business districts.\n\n", options: { color: COLOR_MUTED, fontSize: 9.5 } },
        
        { text: "First/Last Mile Solutions\n", options: { bold: true, color: COLOR_TEXT, fontSize: 11 } },
        { text: "Zero-emission micro-mobility shares integrated with heavy rail lines.", options: { color: COLOR_MUTED, fontSize: 9.5 } }
      ], {
        x: 0.5, y: 1.5, w: 4.5, h: 2.6,
        fontFace: FONT_BODY
      });

      // Mini tags horizontal row
      const tags = ["Clean Air", "Public Space", "Health & Equity"];
      tags.forEach((tag, idx) => {
        const xPos = 0.5 + idx * 1.55;
        slide.addShape('rect', {
          x: xPos, y: 4.25, w: 1.4, h: 0.5,
          fill: CARD_BG, line: { color: ECO_BORDER, width: 1 }
        });
        slide.addText("✔ " + tag, {
          x: xPos, y: 4.35, w: 1.4, h: 0.3,
          fontSize: 8.5, bold: true, color: COLOR_PRIMARY,
          fontFace: FONT_TITLE, align: 'center'
        });
      });

      // Right Column - Transit Image
      slide.addImage({
        path: "images/slide4_mobility.png",
        x: 5.3, y: 1.5, w: 4.2, h: 3.4
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 5: Safe & Affordable Housing
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Safe & Affordable Housing", "Adequate housing is a fundamental human right. Sustainable cities must accommodate all income levels.");

      // Left Column Text
      slide.addText("Rapid, unplanned urbanization has led to 1 billion people living in slums and informal settlements today, lacking basic sanitation and secure tenure.", {
        x: 0.5, y: 1.5, w: 4.5, h: 0.9,
        fontSize: 10, color: COLOR_MUTED,
        fontFace: FONT_BODY, lineSpacing: 14
      });

      // Left Column Focus Areas (3 Cards row)
      const focusAreas = [
        { title: "Participatory Upgrading", desc: "Engaging communities in planning upgrades directly." },
        { title: "Affordable Tech", desc: "Using local, low-carbon modular materials." },
        { title: "Policy & Rights", desc: "Zoning and land tenure to prevent displacement." }
      ];

      focusAreas.forEach((area, idx) => {
        const xPos = 0.5 + idx * 1.55;
        slide.addShape('rect', {
          x: xPos, y: 2.5, w: 1.4, h: 2.4,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        slide.addText(area.title.toUpperCase(), {
          x: xPos + 0.05, y: 2.6, w: 1.3, h: 0.45,
          fontSize: 8, bold: true, color: COLOR_ACCENT,
          fontFace: FONT_TITLE, align: 'center'
        });
        slide.addText(area.desc, {
          x: xPos + 0.05, y: 3.1, w: 1.3, h: 1.7,
          fontSize: 8, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 12, align: 'center'
        });
      });

      // Right Column Housing Image
      slide.addImage({
        path: "images/slide5_housing.png",
        x: 5.3, y: 1.5, w: 4.2, h: 2.2
      });

      // Quote Block below image
      slide.addShape('rect', {
        x: 5.3, y: 3.85, w: 4.2, h: 1.1,
        fill: ACCENT_BG
      });
      slide.addShape('rect', {
        x: 5.3, y: 3.85, w: 0.08, h: 1.1,
        fill: COLOR_ACCENT
      });
      slide.addText("\"Slum upgrading is not just about brick and mortar; it's about integration, dignity, and economic opportunity.\"\n— UN-Habitat Advisory Board", {
        x: 5.5, y: 3.9, w: 3.9, h: 1.0,
        fontSize: 9.5, color: COLOR_TEXT, italic: true,
        fontFace: FONT_BODY, lineSpacing: 14
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 6: Building Climate-Resilient Cities
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Building Climate-Resilient Cities", "As climate change intensifies weather events, cities must shift from reactive recovery to proactive adaptation.");

      // Left Column text/lead
      slide.addText("Mitigating environmental impact and planning disaster resilience is critical to reducing economic and life losses.", {
        x: 0.5, y: 1.4, w: 4.5, h: 0.5,
        fontSize: 10, color: COLOR_MUTED,
        fontFace: FONT_BODY
      });

      // Left Column 3 list items
      const resiliencePoints = [
        { title: "Sponge Cities", desc: "Porous urban surfaces (bioswales, green roofs, permeable concrete) to absorb stormwater runoff." },
        { title: "Early Warning Systems", desc: "Deploying localized sensors and multi-channel notification networks to alert citizens." },
        { title: "Hardened Infrastructure", desc: "Retrofitting seawalls, raising critical utilities, and design of resilient buildings." }
      ];

      resiliencePoints.forEach((point, idx) => {
        const yPos = 2.05 + idx * 1.05;
        // Background strip
        slide.addShape('rect', {
          x: 0.5, y: yPos, w: 4.5, h: 0.95,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Title
        slide.addText(point.title, {
          x: 0.7, y: yPos + 0.1, w: 4.1, h: 0.25,
          fontSize: 11, bold: true, color: COLOR_SECONDARY,
          fontFace: FONT_TITLE
        });
        // Desc
        slide.addText(point.desc, {
          x: 0.7, y: yPos + 0.38, w: 4.1, h: 0.5,
          fontSize: 8.5, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 12
        });
      });

      // Right Column - Resilience Image
      slide.addImage({
        path: "images/slide6_resilience.png",
        x: 5.3, y: 1.5, w: 4.2, h: 3.4
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 7: Zero Waste & Circular Economies
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Zero Waste & Circular Economies", "Transforming waste streams into secondary resource streams to reduce urban footprints.");

      // Left Column: Text + 3 Steps stacked
      slide.addText("Traditional cities follow a linear model: 'Take, Make, Dispose'. Circular cities aim to design waste out of the system entirely, ensuring nutrients flow back into production cycles.", {
        x: 0.5, y: 1.4, w: 4.5, h: 0.8,
        fontSize: 9.5, color: COLOR_MUTED,
        fontFace: FONT_BODY, lineSpacing: 14
      });

      const steps = [
        { num: "1", title: "Source Separation", desc: "Separating organics, dry recyclables, and hazardous wastes at source." },
        { num: "2", title: "Localized Processing", desc: "Anaerobic digesters to produce biogas and fertilizer locally." },
        { num: "3", title: "Upcycling & Repurposing", desc: "One factory's waste is another's raw material in regional parks." }
      ];

      steps.forEach((step, idx) => {
        const yPos = 2.3 + idx * 0.95;
        // Background card
        slide.addShape('rect', {
          x: 0.5, y: yPos, w: 4.5, h: 0.85,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Bubble number
        slide.addShape('oval', {
          x: 0.65, y: yPos + 0.15, w: 0.35, h: 0.35,
          fill: COLOR_SECONDARY
        });
        slide.addText(step.num, {
          x: 0.65, y: yPos + 0.2, w: 0.35, h: 0.25,
          fontSize: 10, bold: true, color: BG_COLOR, align: 'center'
        });
        // Content
        slide.addText(step.title, {
          x: 1.15, y: yPos + 0.08, w: 3.7, h: 0.25,
          fontSize: 10, bold: true, color: COLOR_TEXT,
          fontFace: FONT_TITLE
        });
        slide.addText(step.desc, {
          x: 1.15, y: yPos + 0.35, w: 3.7, h: 0.45,
          fontSize: 8, color: COLOR_MUTED,
          fontFace: FONT_BODY
        });
      });

      // Right Column - Recycling station Image
      slide.addImage({
        path: "images/slide7_waste.png",
        x: 5.3, y: 1.5, w: 4.2, h: 2.1
      });

      // Right Column - Metrics row below image
      // Box 1
      slide.addShape('rect', {
        x: 5.3, y: 3.8, w: 2.0, h: 1.1,
        fill: CARD_BG, line: { color: "252F48", width: 1 }
      });
      slide.addText("80%", {
        x: 5.4, y: 3.9, w: 1.8, h: 0.35,
        fontSize: 18, bold: true, color: COLOR_SECONDARY,
        fontFace: FONT_TITLE, align: 'center'
      });
      slide.addText("RECYCLABLE / COMPOSTABLE", {
        x: 5.4, y: 4.3, w: 1.8, h: 0.5,
        fontSize: 7, bold: true, color: COLOR_MUTED,
        fontFace: FONT_BODY, align: 'center'
      });

      // Box 2
      slide.addShape('rect', {
        x: 7.5, y: 3.8, w: 2.0, h: 1.1,
        fill: CARD_BG, line: { color: "252F48", width: 1 }
      });
      slide.addText("2.2B", {
        x: 7.6, y: 3.9, w: 1.8, h: 0.35,
        fontSize: 18, bold: true, color: COLOR_SECONDARY,
        fontFace: FONT_TITLE, align: 'center'
      });
      slide.addText("ANNUAL SOLID WASTE BY 2025", {
        x: 7.6, y: 4.3, w: 1.8, h: 0.5,
        fontSize: 7, bold: true, color: COLOR_MUTED,
        fontFace: FONT_BODY, align: 'center'
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 8: Reintegrating Nature in Cities
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Reintegrating Nature in Cities", "Expanding green cover to improve public health, support biodiversity, and reduce the Urban Heat Island (UHI) effect.");

      // Left Column bullets
      slide.addText([
        { text: "Micro-Parks & Forests\n", options: { bold: true, color: COLOR_TEXT, fontSize: 11 } },
        { text: "Converting derelict parcels into native pocket forests using Miyawaki methods.\n\n", options: { color: COLOR_MUTED, fontSize: 9.5 } },
        
        { text: "Green Envelopes\n", options: { bold: true, color: COLOR_TEXT, fontSize: 11 } },
        { text: "Mandating green roofs, vertical gardens, and vegetative facade elements.\n\n", options: { color: COLOR_MUTED, fontSize: 9.5 } },
        
        { text: "Ecological Corridors\n", options: { bold: true, color: COLOR_TEXT, fontSize: 11 } },
        { text: "Interlinking public parks to allow wildlife movement and migration.", options: { color: COLOR_MUTED, fontSize: 9.5 } }
      ], {
        x: 0.5, y: 1.5, w: 4.5, h: 2.2,
        fontFace: FONT_BODY
      });

      // Left Column 2 Horizontal Benefits
      const ecoCards = [
        { title: "Urban Cooling Effect", desc: "Lowers surface temps by 2°C to 8°C." },
        { title: "Well-being & Health", desc: "Lower cortisol and cardiovascular risks." }
      ];

      ecoCards.forEach((card, idx) => {
        const xPos = 0.5 + idx * 2.3;
        slide.addShape('rect', {
          x: xPos, y: 3.8, w: 2.2, h: 1.1,
          fill: ECO_BG, line: { color: ECO_BORDER, width: 1 }
        });
        slide.addText("🌱 " + card.title, {
          x: xPos + 0.1, y: 3.9, w: 2.0, h: 0.25,
          fontSize: 9.5, bold: true, color: COLOR_PRIMARY,
          fontFace: FONT_TITLE
        });
        slide.addText(card.desc, {
          x: xPos + 0.1, y: 4.2, w: 2.0, h: 0.65,
          fontSize: 8, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 11
        });
      });

      // Right Column - Green space image
      slide.addImage({
        path: "images/slide8_greenspaces.png",
        x: 5.3, y: 1.5, w: 4.2, h: 3.4
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 9: Smart Infrastructure & Governance
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Smart Infrastructure & Governance", "Using technology not just for surveillance, but to optimize resources and elevate civic participation.");

      // Left Column text
      slide.addText("IoT and data-driven systems optimize resource usage and support citizen governance.", {
        x: 0.5, y: 1.4, w: 4.5, h: 0.5,
        fontSize: 10, color: COLOR_MUTED,
        fontFace: FONT_BODY
      });

      // Left Column 3 list items
      const smartPoints = [
        { title: "Smart Energy Grids", desc: "IoT-enabled grids match power supply to demand, prioritizing local renewables." },
        { title: "Responsive Logistics", desc: "Real-time traffic sensors route transit and optimize waste collections." },
        { title: "Inclusive e-Democracy", desc: "Mobile portals for budgeting and localized hazard reporting by citizens." }
      ];

      smartPoints.forEach((point, idx) => {
        const yPos = 2.05 + idx * 1.05;
        // Background card
        slide.addShape('rect', {
          x: 0.5, y: yPos, w: 4.5, h: 0.95,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Title
        slide.addText("⚡ " + point.title, {
          x: 0.7, y: yPos + 0.1, w: 4.1, h: 0.25,
          fontSize: 11, bold: true, color: COLOR_TEXT,
          fontFace: FONT_TITLE
        });
        // Desc
        slide.addText(point.desc, {
          x: 0.7, y: yPos + 0.38, w: 4.1, h: 0.5,
          fontSize: 8.5, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 12
        });
      });

      // Right Column - Smart City Image
      slide.addImage({
        path: "images/slide9_smartcity.png",
        x: 5.3, y: 1.5, w: 4.2, h: 3.4
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 10: Call to Action & Conclusion
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };

      // Left Accent Ribbon
      slide.addShape('rect', { x: 0, y: 0, w: 0.15, h: 5.625, fill: COLOR_PRIMARY });

      // Title
      slide.addText("A Vision for Tomorrow's Cities", {
        x: 0.6, y: 0.5, w: 4.5, h: 0.5,
        fontSize: 22, bold: true, color: COLOR_TEXT,
        fontFace: FONT_TITLE
      });

      // Subtitle
      slide.addText("Transitioning to sustainable urban habitats requires immediate collaborative action.", {
        x: 0.6, y: 0.95, w: 4.5, h: 0.4,
        fontSize: 10, color: COLOR_MUTED,
        fontFace: FONT_BODY
      });

      // Left Column 3 items
      const actions = [
        { title: "Policy Reform", desc: "Aligning national urban policies with climate and biodiversity targets." },
        { title: "Capital Mobilization", desc: "Leveraging green municipal bonds and adaptation finance." },
        { title: "Citizen Ownership", desc: "Empowering neighborhoods to build and co-manage micro-projects." }
      ];

      actions.forEach((act, idx) => {
        const yPos = 1.5 + idx * 1.2;
        slide.addShape('rect', {
          x: 0.6, y: yPos, w: 4.4, h: 1.05,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        slide.addText(act.title, {
          x: 0.75, y: yPos + 0.1, w: 4.1, h: 0.25,
          fontSize: 11, bold: true, color: COLOR_ACCENT,
          fontFace: FONT_TITLE
        });
        slide.addText(act.desc, {
          x: 0.75, y: yPos + 0.38, w: 4.1, h: 0.6,
          fontSize: 8.5, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 12
        });
      });

      // Right Column - Vision Image
      slide.addImage({
        path: "images/slide10_vision.png",
        x: 5.3, y: 1.0, w: 4.2, h: 2.6
      });

      // Divider Line
      slide.addShape('rect', { x: 5.3, y: 3.8, w: 4.2, h: 0.02, fill: COLOR_PRIMARY });

      // Closing Quote
      slide.addText("\"The battle for sustainability will be won or lost in our cities.\"  — UN-Habitat", {
        x: 5.3, y: 4.0, w: 4.2, h: 1.0,
        fontSize: 11, bold: true, color: COLOR_PRIMARY, align: 'center', italic: true,
        fontFace: FONT_TITLE
      });
    }

    // ------------------------------------------------------------------------
    // 3. Export Presentation
    // ------------------------------------------------------------------------
    pptx.writeFile({ fileName: "Sustainable_Cities_and_Communities_SDG11.pptx" })
      .then(() => {
        console.log("PPTX generation complete!");
      })
      .catch(err => {
        console.error("Error writing PPTX:", err);
        alert("Failed to export PowerPoint presentation. See console for details.");
      })
      .finally(() => {
        // Reset download button
        downloadBtn.querySelector('span').innerHTML = originalText;
        downloadBtn.disabled = false;
      });

  } catch (error) {
    console.error("PPTX generation error:", error);
    alert("Error occurred while generating PowerPoint: " + error.message);
    downloadBtn.querySelector('span').innerHTML = originalText;
    downloadBtn.disabled = false;
  }
}

// Kickoff
document.addEventListener('DOMContentLoaded', initPresentation);
