export type Question = {
  key: string;
  commentKey: string;
  text: string;
};

export const CATEGORIES: { name: string; questions: Question[] }[] = [
  {
    name: "Position Specific Competencies and Productivity",
    questions: [
      { key: "PSC_DutiesScale", commentKey: "PSC_DutiesComment", text: "Performs all essential duties and functions associated with the position." },
      { key: "PSC_Expected", commentKey: "PSC_ExpectedComment", text: "Produces the amount of work expected of the position." },
      { key: "PSC_Timelines", commentKey: "PSC_TimelinesComment", text: "Works within timelines to meet individual and team goals." },
      { key: "PSC_SeeksTasks", commentKey: "PSC_SeeksTasksComment", text: "Proactively seeks out new tasks, opportunities, and/or projects to work on." },
      { key: "PSC_IncreasedRespo", commentKey: "PSC_IncreasedRespoComment", text: "Actively seeks increased responsibility within the team and/or organization." }
    ]
  },
  {
    name: "Responsibility",
    questions: [
      { key: "RES_Deadlines", commentKey: "RES_DeadlinesComment", text: "Consistently follows through on commitments and meets deadlines." },
      { key: "RES_AddrMistakes", commentKey: "RES_AddrMistakesComment", text: "Is accountable for and addresses mistakes promptly." },
      { key: "RES_Independent", commentKey: "RES_IndependentComment", text: "Works well independently, staying self-motivated and self-directed." },
      { key: "RES_Training", commentKey: "RES_TrainingComment", text: "Proactively seeks opportunities for training." }
    ]
  },
  {
    name: "Task Management",
    questions: [
      { key: "TSK_Priori", commentKey: "TSK_PrioriComment", text: "Organized and skilled at prioritizing tasks." },
      { key: "TSK_AllocatesTime", commentKey: "TSK_AllocatesTimeComment", text: "Allocates the time needed to complete work." },
      { key: "TSK_Strategies", commentKey: "TSK_StrategiesComment", text: "Evaluates and implements effective strategies to save time when able." },
      { key: "TSK_UnderPress", commentKey: "TSK_UnderPressComment", text: "Works well under pressure without sacrificing the quality of work." },
      { key: "TSK_NeatOrganized", commentKey: "TSK_NeatOrganizedComment", text: "Contributes to a neat and organized work environment." }
    ]
  },
  {
    name: "Critical Thinking Skills",
    questions: [
      { key: "CTS_Angles", commentKey: "CTS_AnglesComment", text: "Analyzes challenges and opportunities from various angles and perspectives." },
      { key: "CTS_Change", commentKey: "CTS_ChangeComment", text: "Navigates change and/or shifting priorities with openness and flexibility." },
      { key: "CTS_Assistance", commentKey: "CTS_AssistanceComment", text: "Asks for assistance when needed." }
    ]
  },
  {
    name: "Communication & Cooperation",
    questions: [
      { key: "CC_VerbalWr", commentKey: "CC_VerbalWrComment", text: "Is effective in both verbal and written communication to keep others informed." },
      { key: "CC_AcptFeedback", commentKey: "CC_AcptFeedbackComment", text: "Accepts feedback that improves their own performance." },
      { key: "CC_DelivFeedback", commentKey: "CC_DelivFeedbackComment", text: "Delivers feedback that improves the performance of others." }
    ]
  },
  {
    name: "Interpersonal Functions",
    questions: [
      { key: "IF_Clients", commentKey: "IF_ClientsComment", text: "Establishes strong relationships with clients." },
      { key: "IF_Colleagues", commentKey: "IF_ColleaguesComment", text: "Builds collaborative relationships with colleagues." },
      { key: "IF_Motivates", commentKey: "IF_MotivatesComment", text: "Motivates and encourages others." },
      { key: "IF_Complaints", commentKey: "IF_ComplaintsComment", text: "Resolves client complaints professionally and promptly." }
    ]
  }
];
