const verbatimQuotesData = {
  title: "Do you want to explain any of your previous answers?",
  section: "VERBATIM QUOTES",
  company: "Ericsson",
  quotes: [
    "I rated some of the development questions lower because I don’t see Marcus coaching people directly very often. He may be doing it with his own team, but I don’t have enough exposure to judge that fully.",

    "My lower ratings on performance standards are not because I think Marcus lacks standards. It is more that I don’t always see the follow-through after priorities are set.",

    "I scored him highly on collaboration because he is very good at involving the right people before decisions are made. That makes a big difference.",

    "On stopping low-value work, I think there is room to improve. There are still projects that continue because they have already started, even when the value is unclear.",

    "I may not have enough visibility into how he develops his direct reports, so my scores there are based mainly on what I observe in broader meetings.",

    "Marcus is strong at setting direction, but the direction can change quickly. When that happens, it would help to be clearer on what is no longer a priority.",

    "I gave lower scores on coaching because feedback usually happens when someone asks for it. I don’t often see him proactively giving development input."
  ]
};
export default verbatimQuotesData;
export const positiveImpactData = {
  title:
    "What is one thing this leader does that makes the biggest positive impact?",
    
  section: "VERBATIM QUOTES",

  company: "Ericsson",

  quotes: [
    "Marcus is very good at getting the right people into the discussion early. It means we avoid a lot of rework later.",

    "He is one of the few leaders who can cut across functions without making it political.",

    "When there is a blocker between teams, Marcus usually knows who needs to be involved and gets them aligned quickly.",

    "I’ve seen him bring Finance, Product and Commercial together in a way that helped us make decisions faster.",

    "He creates enough space for people to disagree, but usually keeps the conversation moving.",

    "Marcus is trusted across teams. People will take his call even when the topic is difficult.",

    "He is good at spotting when a decision will fail unless another function has been properly included.",

    "I think his biggest strength is that he does not just solve things within his own area — he thinks about the wider system.",

    "He brings calm to cross-functional discussions, especially when there are competing priorities.",

    "Marcus is clear on what matters. I usually leave meetings knowing what we agreed and who owns what."
  ]
};
export const leadershipImpactData = {
  title:
    "What’s one change this person could make to increase their leadership impact?",

  section: "VERBATIM QUOTES",

  company: "Ericsson",

  quotes: [
    "Marcus is supportive, but I would like more specific development feedback from him.",

    "My 1:1s are useful, but they are more focused on work delivery than on my growth.",

    "He gives encouragement, but not always enough concrete advice on what I should do differently.",

    "I think he sees people’s potential, but he does not always turn that into a clear development plan.",

    "Feedback tends to come when I ask for it rather than proactively.",

    "He is generous with time when there is an issue, but less structured on longer-term career development.",

    "I would value more coaching from him on how to operate at the next level.",

    "He connects the work to the business context, but not always to individual development goals.",

    "Marcus could be more deliberate about stretching people and explaining what they need to learn from the stretch."
  ]
};
export const personalDevelopmentPlanData = {
  title: "Personal development plan",

  section: "DRAFT FOR INPUT",

  company: "Ericsson",

  professionalGoals: {
    title:
      "Professional Goals What outcomes do you want to deliver over the next 12 months?",

    withinBU: [
      {
        text: "[Examples]"
      },
      {
        text: "[Examples]"
      },
      {
        text: "[Examples]"
      }
    ],

    acrossOrganization: [
      {
        text: "[Examples]"
      },
      {
        text: "[Examples]"
      },
      {
        text: "[Examples]"
      }
    ]
  },

  behaviorDevelopment: {
    title:
      "Behavior Development Pick two or three behaviors to build. What does success look like, and how will you know?",

    headers: [
      {
        text: "Behavior",
        options: { italic: true }
      },
      {
        text: "Success Definition",
        options: { italic: true }
      },
      {
        text: "Measurable Outcomes",
        options: { italic: true }
      }
    ],

    rows: [
      [
        {
          text: "[Behavior]"
        },
        {
          text: "[Example]"
        },
        {
          text: "• [Example]"
        }
      ],

      [
        {
          text: "[behavior]"
        },
        {
          text: "[Example]"
        },
        {
          text: "• [Example]"
        }
      ],

      [
        {
          text: "[behavior]"
        },
        {
          text: "[Example]"
        },
        {
          text: "• [Example]"
        }
      ]
    ]
  },

  actionPlan: {
    title:
      "Action Plan What will you do in the next 30 days to start building these behaviors?",

    items: [
      {
        text: "1. [Example]"
      }
    ]
  },

  stakeholders: {
    title:
      "Stakeholders & Support Who do you need to engage, and what help will you ask for?",

    items: [
      {
        text: "1. [Example]"
      }
    ]
  }
};

export const verbatimThemesData = {
  title: "Verbatim Quotes",

  section: "MANUAL REPORT ONLY",

  company: "Ericsson",

  themes: [
    {
      title: "Cross-functional impact",
      rating: "7 of 12 raters",
      headerColor: "16A34A",

      linkedBehaviors:
        "Linked micro-behaviors: Involves relevant stakeholders early · Encourages open challenge",

      quotes: [
        "“Marcus is the person I call first when I need something unblocked across teams. He's genuinely trusted.”",

        "“He brings the right people in early. Decisions stick because everyone has been in the room.”"
      ]
    },

    {
      title: "Direction & clarity",
      rating: "6 of 12 raters",
      headerColor: "16A34A",

      linkedBehaviors:
        "Linked micro-behaviors: Sets direction and priorities with clear owners · Removes obstacles to progress",

      quotes: [
        "“You always know where you stand with Marcus - goals are crystal clear from day one.”",

        "“When things get noisy, his framing brings the temperature down and gets us moving.”"
      ]
    },

    {
      title: "Coaching & development",
      rating: "5 of 12 raters",
      headerColor: "F26A3D",

      linkedBehaviors:
        "Linked micro-behaviors: Develops others for long-term impact · Connects work to strategy and customers",

      quotes: [
        "“I leave 1:1s feeling supported but without a concrete development action.”",

        "“Marcus is supportive, but I have to ask for feedback to get it. It doesn’t come unprompted.”"
      ]
    },

    {
      title: "Outcome discipline",
      rating: "4 of 12 raters",
      headerColor: "F26A3D",

      linkedBehaviors:
        "Linked micro-behaviors: Sets high standards for quality and performance · Stops work that no longer adds value",

      quotes: [
        "“He sets the vision well, but follow-through on metrics is patchy - I rarely see hard numbers tracked back.”",

        "“Priorities can shift week to week. I’d value a clearer signal on what to drop.”"
      ]
    }
  ],

  footerNote:
    "A full quote book is provided separately. Themes here represent a curated subset; counts reflect raters who touched on each theme."
};

export const coachingQuadrantData = {
  title: "360 results: Coaching quadrant",

  section: "MANUAL REPORT ONLY",

  company: "Ericsson",

  quadrants: [
    {
      title: "Under-recognized strengths",

      subtitle:
        "Colleagues see these more consistently than you do.",

      items: [
        "Seeks input from affected stakeholders before moving forward"
      ],

      bgColor: "D9D9D9"
    },

    {
      title: "Shared development areas",

      subtitle:
        "You and colleagues both see these as less consistent.",

      items: [
        "Gives others ownership within boundaries"
      ],

      bgColor: "F2F2F2"
    },

    {
      title: "Shared strengths",

      subtitle:
        "You think you show these often, and colleagues broadly agree.",

      items: [
        "Sets direction and priorities",
        "Explains decision trade-offs",
        "Removes obstacles to progress"
      ],

      bgColor: "F2F2F2"
    },

    {
      title: "Potential blind spots",

      subtitle:
        "You think you show these often, but colleagues do not see them consistently.",

      items: [
        "Sets high standards for performance",
        "Develops others for long-term impact",
        "Stops work that no longer adds value"
      ],

      bgColor: "D9D9D9"
    }
  ]
};
export const resultsAtGlanceData = {
  title: "Your results at a glance",

  subtitle:
    "Executive summary of the key points from your 360 diagnostic report",

  sectionLeft: "Micro-behaviors not yet finalized",

  sectionRight: "‘What to do’ pending Ericsson input",

  company: "Ericsson",

  participantInfo: {
    name: "Marcus Gosstner",
    reportingPeriod: "May 2026",
    responseRate: "13 / 15",

    confidentialityText:
      "This report is confidential to Ericsson and is intended solely for leadership development purposes. The feedback included in this report has been gathered from selected colleagues and summarized to provide a balanced view of the participant’s leadership behaviors, strengths, and development areas.",

    anonymityText:
      "Individual respondent feedback has not been attributed by name. Results are presented in aggregate wherever possible to protect respondent anonymity. The report should not be shared beyond the agreed audience without prior approval."
  },

  feedbackSections: [
    {
      title:
        "Your most consistent leadership micro-behaviors are",

      points: [
        "Sets direction and priorities",
        "Seeks input from affected stakeholders before moving forward"
      ],

      action:
        "Lean into these strengths and make a conscious effort to take on work that complements them."
    },

    {
      title:
        "Your least consistent leadership micro-behaviors are",

      points: [
        "Sets high standards for performance",
        "Stops work that no longer adds value",
        "Develops others for long-term impact"
      ],

      action:
        "Build 2–3 into your development plan. Define what change you’d like to see and proactively seek feedback on your development."
    },

    {
      title: "Your potential blind spots are",

      points: [
        "Develops others for long-term impact",
        "Connects work to vision and strategy",
        "Sets high standards for quality and performance"
      ],

      action:
        "Reflect on the perception gap.\n\nWhere you over-rate yourself, test intent against impact.\n\nWhere you under-rate yourself, focus on deliberate visibility."
    },

    {
      title:
        "Your potential under-recognized strengths are",

      points: [
        "Seeks input from affected stakeholders before moving forward",
        "Sets direction and priorities with clear owners",
        "Facilitates collective momentum across teams"
      ],

      action:
        "We encourage you to solicit feedback on your development journey."
    }
  ]
};