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
        fontSize: 22, bold: true, color: COLOR_TEXT,
        fontFace: FONT_TITLE
      });

      // Add Subtitle / Lead text if provided
      if (subtitleText) {
        slide.addText(subtitleText, {
          x: 0.5, y: 0.95, w: 9.0, h: 0.35,
          fontSize: 11, color: COLOR_MUTED,
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
        x: 0.8, y: 1.2, w: 8.4, h: 0.3,
        fontSize: 10, bold: true, color: COLOR_PRIMARY,
        fontFace: FONT_TITLE, characterSpacing: 1
      });

      // Title
      slide.addText("Sustainable Cities &\nCommunities", {
        x: 0.8, y: 1.6, w: 8.4, h: 1.4,
        fontSize: 38, bold: true, color: COLOR_TEXT,
        fontFace: FONT_TITLE
      });

      // Subtitle
      slide.addText("Shaping the Urban Future through Innovation, Ecology, and Equity", {
        x: 0.8, y: 3.1, w: 8.4, h: 0.7,
        fontSize: 15, color: COLOR_MUTED,
        fontFace: FONT_BODY
      });

      // Horizontal Divider Line
      slide.addShape('rect', { x: 0.8, y: 3.9, w: 2.0, h: 0.03, fill: COLOR_PRIMARY });

      // Footer Meta
      slide.addText("July 2026  |  Global Development Agenda", {
        x: 0.8, y: 4.1, w: 8.4, h: 0.3,
        fontSize: 9.5, color: COLOR_ACCENT,
        fontFace: FONT_BODY
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 2: The Urban Surge
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "The Urban Surge", "Understanding the massive scale of modern global urbanization.");

      // Left Column Content
      slide.addText("Today, more than half of the world's population lives in cities. By 2050, that figure is projected to rise to nearly 70%.", {
        x: 0.5, y: 1.6, w: 4.5, h: 1.1,
        fontSize: 13, bold: true, color: COLOR_TEXT,
        fontFace: FONT_BODY, lineSpacing: 18
      });

      slide.addText("Cities are the engines of economic growth, accounting for about 80% of global GDP. However, they also consume over two-thirds of the world's energy and account for more than 70% of global carbon emissions.\n\nThe future of sustainable global development relies heavily on how we design and manage urban centers.", {
        x: 0.5, y: 2.8, w: 4.5, h: 2.2,
        fontSize: 11, color: COLOR_MUTED,
        fontFace: FONT_BODY, lineSpacing: 16
      });

      // Right Column Content - Stats Stack
      const stats = [
        { num: "4.4 Billion", label: "People live in cities today" },
        { num: "70%", label: "Of global CO2 emissions from cities" },
        { num: "60%", label: "Projected urban population by 2030" }
      ];

      stats.forEach((stat, idx) => {
        const yPos = 1.5 + idx * 1.25;
        // Background card shape
        slide.addShape('rect', {
          x: 5.3, y: yPos, w: 4.2, h: 1.05,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Stat number
        slide.addText(stat.num, {
          x: 5.5, y: yPos + 0.12, w: 3.8, h: 0.45,
          fontSize: 22, bold: true, color: COLOR_PRIMARY,
          fontFace: FONT_TITLE
        });
        // Stat label
        slide.addText(stat.label.toUpperCase(), {
          x: 5.5, y: yPos + 0.58, w: 3.8, h: 0.3,
          fontSize: 8.5, bold: true, color: COLOR_MUTED,
          fontFace: FONT_BODY
        });
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 3: Core Pillars of SDG 11
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Core Pillars of Sustainable Urbanism", "Achieving SDG 11 requires a holistic approach addressing multiple interconnected systems.");

      const pillars = [
        {
          title: "Housing & Services",
          desc: "Ensuring access to safe, affordable, and basic housing, plus upgrading slum settlements globally."
        },
        {
          title: "Sustainable Mobility",
          desc: "Expanding safe, affordable, and accessible public transit systems with emphasis on active travel."
        },
        {
          title: "Urban Resilience",
          desc: "Improving disaster preparedness, mitigating environmental impact, and reducing loss of life."
        },
        {
          title: "Circular Flow",
          desc: "Reducing municipal waste, improving air quality, and supporting localized green economies."
        }
      ];

      const colWidth = 2.1;
      const gap = 0.2;
      const startX = 0.5;

      pillars.forEach((pillar, idx) => {
        const xPos = startX + idx * (colWidth + gap);
        // Card Background
        slide.addShape('rect', {
          x: xPos, y: 1.6, w: colWidth, h: 3.4,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Title
        slide.addText(pillar.title, {
          x: xPos + 0.15, y: 1.8, w: colWidth - 0.3, h: 0.6,
          fontSize: 13, bold: true, color: COLOR_SECONDARY,
          fontFace: FONT_TITLE
        });
        // Description
        slide.addText(pillar.desc, {
          x: xPos + 0.15, y: 2.5, w: colWidth - 0.3, h: 2.3,
          fontSize: 9.5, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 14
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

      // Left Column - Bullet-like blocks
      slide.addText([
        { text: "Integrated Mass Transit\n", options: { bold: true, color: COLOR_TEXT, fontSize: 12 } },
        { text: "High-capacity rapid transit networks powered by renewable electricity to connect far-reaching suburban communities.\n\n", options: { color: COLOR_MUTED, fontSize: 10 } },
        
        { text: "Active Infrastructure\n", options: { bold: true, color: COLOR_TEXT, fontSize: 12 } },
        { text: "Protected cycling networks and fully pedestrianized central business districts to promote human-powered travel.\n\n", options: { color: COLOR_MUTED, fontSize: 10 } },
        
        { text: "First/Last Mile Solutions\n", options: { bold: true, color: COLOR_TEXT, fontSize: 12 } },
        { text: "Zero-emission micro-mobility shares integrated with heavy rail lines to solve short journey bottlenecks.", options: { color: COLOR_MUTED, fontSize: 10 } }
      ], {
        x: 0.5, y: 1.6, w: 4.5, h: 3.4,
        fontFace: FONT_BODY
      });

      // Right Column - Benefits list
      const benefits = [
        { title: "Reduce Air Pollution", desc: "Eradicates vehicle emissions, directly improving local air quality indices." },
        { title: "Reclaim Public Space", desc: "Replaces massive parking blocks and highways with green parks and plazas." },
        { title: "Enhance Health & Equity", desc: "Encourages daily activity and ensures low-cost transport accessibility." }
      ];

      benefits.forEach((benefit, idx) => {
        const yPos = 1.6 + idx * 1.15;
        // Background card
        slide.addShape('rect', {
          x: 5.3, y: yPos, w: 4.2, h: 0.95,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Checkmark Icon indicator
        slide.addText("✔", {
          x: 5.45, y: yPos + 0.15, w: 0.3, h: 0.3,
          fontSize: 14, color: COLOR_PRIMARY, bold: true
        });
        // Title
        slide.addText(benefit.title, {
          x: 5.8, y: yPos + 0.12, w: 3.5, h: 0.25,
          fontSize: 11, bold: true, color: COLOR_PRIMARY,
          fontFace: FONT_TITLE
        });
        // Desc
        slide.addText(benefit.desc, {
          x: 5.8, y: yPos + 0.42, w: 3.5, h: 0.45,
          fontSize: 9, color: COLOR_MUTED,
          fontFace: FONT_BODY
        });
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 5: Safe & Affordable Housing
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Safe & Affordable Housing", "Adequate housing is a fundamental human right. Sustainable cities must accommodate all income levels.");

      // Left Column: Text + Indigo Quote Card
      slide.addText("Rapid, unplanned urbanization has led to 1 billion people living in slums and informal settlements today, lacking basic sanitation, clean water, and secure tenure.", {
        x: 0.5, y: 1.5, w: 4.5, h: 1.3,
        fontSize: 11, color: COLOR_MUTED,
        fontFace: FONT_BODY, lineSpacing: 16
      });

      // Quote Block
      slide.addShape('rect', {
        x: 0.5, y: 2.9, w: 4.5, h: 1.9,
        fill: ACCENT_BG
      });
      // Left border of quote
      slide.addShape('rect', {
        x: 0.5, y: 2.9, w: 0.08, h: 1.9,
        fill: COLOR_ACCENT
      });
      slide.addText("\"Slum upgrading is not just about brick and mortar; it's about integration, dignity, and economic opportunity.\"\n\n— UN-Habitat Advisory Board", {
        x: 0.75, y: 3.1, w: 4.1, h: 1.5,
        fontSize: 11, color: COLOR_TEXT, italic: true,
        fontFace: FONT_BODY, lineSpacing: 16
      });

      // Right Column: Stacked Cards
      const focusAreas = [
        { title: "Participatory Upgrading", desc: "Engaging communities in planning upgrades, ensuring they aren't displaced from their social fabrics." },
        { title: "Affordable Building Tech", desc: "Utilizing local, low-carbon materials and modular building methods to reduce housing costs." },
        { title: "Policy & Land Rights", desc: "Implementing progressive zoning and secure land tenure to prevent sudden displacements." }
      ];

      focusAreas.forEach((area, idx) => {
        const yPos = 1.5 + idx * 1.15;
        slide.addShape('rect', {
          x: 5.3, y: yPos, w: 4.2, h: 1.0,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        slide.addText("FOCUS AREA", {
          x: 5.5, y: yPos + 0.1, w: 3.8, h: 0.2,
          fontSize: 7.5, bold: true, color: COLOR_ACCENT,
          fontFace: FONT_TITLE
        });
        slide.addText(area.title, {
          x: 5.5, y: yPos + 0.25, w: 3.8, h: 0.25,
          fontSize: 11, bold: true, color: COLOR_TEXT,
          fontFace: FONT_TITLE
        });
        slide.addText(area.desc, {
          x: 5.5, y: yPos + 0.52, w: 3.8, h: 0.4,
          fontSize: 9, color: COLOR_MUTED,
          fontFace: FONT_BODY
        });
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 6: Building Climate-Resilient Cities
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Building Climate-Resilient Cities", "As climate change intensifies weather events, cities must shift from reactive recovery to proactive adaptation.");

      const resilientPoints = [
        {
          title: "Sponge Cities",
          desc: "Designing porous urban surfaces (bioswales, green roofs, permeable concrete) to absorb stormwater runoff and recharge groundwater."
        },
        {
          title: "Early Warning Systems",
          desc: "Deploying localized sensors and multi-channel notification networks to alert citizens before disaster strikes."
        },
        {
          title: "Hardened Infrastructure",
          desc: "Retrofitting seawalls, raising critical utility lines, and designing buildings to withstand seismic and extreme meteorological events."
        }
      ];

      const colWidth = 2.8;
      const gap = 0.3;
      const startX = 0.5;

      resilientPoints.forEach((point, idx) => {
        const xPos = startX + idx * (colWidth + gap);
        // Card Background
        slide.addShape('rect', {
          x: xPos, y: 1.7, w: colWidth, h: 3.2,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Title
        slide.addText(point.title, {
          x: xPos + 0.15, y: 2.0, w: colWidth - 0.3, h: 0.5,
          fontSize: 14, bold: true, color: COLOR_SECONDARY,
          fontFace: FONT_TITLE
        });
        // Description
        slide.addText(point.desc, {
          x: xPos + 0.15, y: 2.6, w: colWidth - 0.3, h: 2.1,
          fontSize: 9.5, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 15
        });
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 7: Zero Waste & Circular Economies
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Zero Waste & Circular Economies", "Transforming waste streams into secondary resource streams to reduce urban footprints.");

      // Left Column: Text + 2 Metric Boxes
      slide.addText("Traditional cities follow a linear model: 'Take, Make, Dispose'. Circular cities aim to design waste out of the system entirely, ensuring nutrients and materials flow back into production cycles.", {
        x: 0.5, y: 1.6, w: 4.5, h: 1.3,
        fontSize: 11, color: COLOR_MUTED,
        fontFace: FONT_BODY, lineSpacing: 16
      });

      // Metric Box 1
      slide.addShape('rect', {
        x: 0.5, y: 3.1, w: 2.1, h: 1.7,
        fill: CARD_BG, line: { color: "252F48", width: 1 }
      });
      slide.addText("80%", {
        x: 0.65, y: 3.3, w: 1.8, h: 0.45,
        fontSize: 26, bold: true, color: COLOR_SECONDARY,
        fontFace: FONT_TITLE
      });
      slide.addText("OF URBAN WASTE CAN BE RECYCLED OR COMPOSTED", {
        x: 0.65, y: 3.8, w: 1.8, h: 0.8,
        fontSize: 8, bold: true, color: COLOR_MUTED,
        fontFace: FONT_BODY
      });

      // Metric Box 2
      slide.addShape('rect', {
        x: 2.9, y: 3.1, w: 2.1, h: 1.7,
        fill: CARD_BG, line: { color: "252F48", width: 1 }
      });
      slide.addText("2.2B", {
        x: 3.05, y: 3.3, w: 1.8, h: 0.45,
        fontSize: 26, bold: true, color: COLOR_SECONDARY,
        fontFace: FONT_TITLE
      });
      slide.addText("TONS OF MUNICIPAL SOLID WASTE ANNUALLY BY 2025", {
        x: 3.05, y: 3.8, w: 1.8, h: 0.8,
        fontSize: 8, bold: true, color: COLOR_MUTED,
        fontFace: FONT_BODY
      });

      // Right Column: 3 Steps
      const steps = [
        { num: "1", title: "Source Separation", desc: "Mandating the separation of organics, dry recyclables, and hazardous wastes at source." },
        { num: "2", title: "Localized Processing", desc: "Establishing anaerobic digesters for organic waste to produce biogas and fertilizer." },
        { num: "3", title: "Upcycling & Repurposing", desc: "Creating regional industrial parks where one factory's waste is another's raw material." }
      ];

      steps.forEach((step, idx) => {
        const yPos = 1.6 + idx * 1.1;
        // Background card
        slide.addShape('rect', {
          x: 5.3, y: yPos, w: 4.2, h: 0.95,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Step bubble
        slide.addShape('oval', {
          x: 5.5, y: yPos + 0.15, w: 0.45, h: 0.45,
          fill: COLOR_SECONDARY
        });
        slide.addText(step.num, {
          x: 5.5, y: yPos + 0.22, w: 0.45, h: 0.3,
          fontSize: 12, bold: true, color: BG_COLOR, align: 'center'
        });
        // Step Content
        slide.addText(step.title, {
          x: 6.1, y: yPos + 0.12, w: 3.2, h: 0.25,
          fontSize: 11, bold: true, color: COLOR_TEXT,
          fontFace: FONT_TITLE
        });
        slide.addText(step.desc, {
          x: 6.1, y: yPos + 0.38, w: 3.2, h: 0.45,
          fontSize: 8.5, color: COLOR_MUTED,
          fontFace: FONT_BODY
        });
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 8: Reintegrating Nature in Cities
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Reintegrating Nature in Cities", "Expanding green cover to improve public health, support biodiversity, and reduce the Urban Heat Island (UHI) effect.");

      // Left Column: Text + bullet points
      slide.addText([
        { text: "Micro-Parks & Forests\n", options: { bold: true, color: COLOR_TEXT, fontSize: 12 } },
        { text: "Converting derelict parcels into high-density native pocket forests using Miyawaki methods.\n\n", options: { color: COLOR_MUTED, fontSize: 10 } },
        
        { text: "Green Envelopes\n", options: { bold: true, color: COLOR_TEXT, fontSize: 12 } },
        { text: "Mandating green roofs, vertical gardens, and vegetative facade elements on new developments.\n\n", options: { color: COLOR_MUTED, fontSize: 10 } },
        
        { text: "Ecological Corridors\n", options: { bold: true, color: COLOR_TEXT, fontSize: 12 } },
        { text: "Interlinking public parks to allow wildlife movement and migration pathways.", options: { color: COLOR_MUTED, fontSize: 10 } }
      ], {
        x: 0.5, y: 1.6, w: 4.5, h: 3.4,
        fontFace: FONT_BODY
      });

      // Right Column: Two ecological cards (designed in forest-green theme)
      const ecoCards = [
        {
          title: "Urban Cooling Effect",
          desc: "Trees and vegetation lower surface temperatures by up to 2°C to 8°C through shading and evapotranspiration, mitigating heat risks."
        },
        {
          title: "Mental & Physical Well-being",
          desc: "Access to public green space is proven to lower cortisol levels, reduce cardiovascular risks, and foster social cohesion."
        }
      ];

      ecoCards.forEach((card, idx) => {
        const yPos = 1.6 + idx * 1.7;
        slide.addShape('rect', {
          x: 5.3, y: yPos, w: 4.2, h: 1.5,
          fill: ECO_BG, line: { color: ECO_BORDER, width: 1 }
        });
        slide.addText("🌱  " + card.title, {
          x: 5.5, y: yPos + 0.15, w: 3.8, h: 0.3,
          fontSize: 12, bold: true, color: COLOR_PRIMARY,
          fontFace: FONT_TITLE
        });
        slide.addText(card.desc, {
          x: 5.5, y: yPos + 0.5, w: 3.8, h: 0.85,
          fontSize: 9.5, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 14
        });
      });
    }

    // ------------------------------------------------------------------------
    // SLIDE 9: Smart Infrastructure & Governance
    // ------------------------------------------------------------------------
    {
      const slide = pptx.addSlide();
      slide.background = { fill: BG_COLOR };
      addSlideHeader(slide, "Smart Infrastructure & Governance", "Using technology not just for surveillance, but to optimize resources and elevate civic participation.");

      const smartPoints = [
        {
          title: "Smart Energy Grids",
          desc: "IoT-enabled grids that match power supply to real-time demand, prioritizing local renewable generation and cutting energy wastage."
        },
        {
          title: "Responsive Logistics",
          desc: "Real-time traffic sensors that route transit dynamically and schedule municipal waste pickup only when containers are full."
        },
        {
          title: "Inclusive e-Democracy",
          desc: "Mobile portals for participatory budgeting and localized hazard reporting, ensuring citizen voices directly drive urban projects."
        }
      ];

      const colWidth = 2.8;
      const gap = 0.3;
      const startX = 0.5;

      smartPoints.forEach((point, idx) => {
        const xPos = startX + idx * (colWidth + gap);
        // Card Background
        slide.addShape('rect', {
          x: xPos, y: 1.7, w: colWidth, h: 3.2,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        // Icon Indicator
        slide.addText("⚡", {
          x: xPos + 0.15, y: 1.85, w: 0.4, h: 0.3,
          fontSize: 14, color: COLOR_PRIMARY
        });
        // Title
        slide.addText(point.title, {
          x: xPos + 0.15, y: 2.2, w: colWidth - 0.3, h: 0.5,
          fontSize: 13, bold: true, color: COLOR_TEXT,
          fontFace: FONT_TITLE
        });
        // Description
        slide.addText(point.desc, {
          x: xPos + 0.15, y: 2.7, w: colWidth - 0.3, h: 2.0,
          fontSize: 9.5, color: COLOR_MUTED,
          fontFace: FONT_BODY, lineSpacing: 14
        });
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
        x: 1.0, y: 0.6, w: 8.0, h: 0.5,
        fontSize: 24, bold: true, color: COLOR_TEXT, align: 'center',
        fontFace: FONT_TITLE
      });

      // Subtitle
      slide.addText("Transitioning to sustainable urban habitats requires immediate, collaborative effort across all sectors.", {
        x: 1.0, y: 1.05, w: 8.0, h: 0.4,
        fontSize: 11, color: COLOR_MUTED, align: 'center',
        fontFace: FONT_BODY
      });

      // 3 Action boxes
      const actions = [
        { title: "Policy Reform", desc: "Aligning national urban policies with climate and biodiversity targets." },
        { title: "Capital Mobilization", desc: "Leveraging green municipal bonds and global adaptation finance." },
        { title: "Citizen Ownership", desc: "Empowering neighborhoods to build and co-manage micro-projects." }
      ];

      const colWidth = 2.8;
      const gap = 0.3;
      const startX = 0.5;

      actions.forEach((act, idx) => {
        const xPos = startX + idx * (colWidth + gap);
        slide.addShape('rect', {
          x: xPos, y: 1.7, w: colWidth, h: 2.0,
          fill: CARD_BG, line: { color: "252F48", width: 1 }
        });
        slide.addText(act.title, {
          x: xPos + 0.15, y: 1.9, w: colWidth - 0.3, h: 0.3,
          fontSize: 12, bold: true, color: COLOR_ACCENT, align: 'center',
          fontFace: FONT_TITLE
        });
        slide.addText(act.desc, {
          x: xPos + 0.15, y: 2.3, w: colWidth - 0.3, h: 1.2,
          fontSize: 9, color: COLOR_MUTED, align: 'center',
          fontFace: FONT_BODY, lineSpacing: 13
        });
      });

      // Divider Line
      slide.addShape('rect', { x: 3.0, y: 4.1, w: 4.0, h: 0.02, fill: COLOR_PRIMARY });

      // Closing Quote
      slide.addText("\"The battle for sustainability will be won or lost in our cities.\"  — UN-Habitat", {
        x: 1.0, y: 4.35, w: 8.0, h: 0.5,
        fontSize: 13, bold: true, color: COLOR_PRIMARY, align: 'center', italic: true,
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
