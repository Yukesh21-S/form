import "dotenv/config";
import { PrismaClient } from "@prisma/client";
import { PrismaPg } from "@prisma/adapter-pg";

const adapter = new PrismaPg(
  process.env.DATABASE_URL ?? "postgresql://postgres:123@localhost:5432/form"
);
const prisma = new PrismaClient({ adapter });

const OPTIONS = [
  { label: "Rarely",                value: 25   },
  { label: "Sometimes",             value: 50   },
  { label: "Often",                 value: 75   },
  { label: "Always",                value: 100  },
  { label: "Insufficient Exposure", value: null },
] as const;

const QUESTIONS = [
  { text: "Takes ownership of outcomes end-to-end" },
  { text: "Prioritizes highest-value work clearly" },
  { text: "Makes timely decisions under uncertainty" },
  { text: "Questions assumptions and ways of working" },
  { text: "Adapts based on feedback" },
  { text: "Runs continuous experiments instead of waiting" },
  { text: "Uses external signals to guide decisions" },
  { text: "Proactively seeks feedback to improve impact" },
  { text: "Follows through on commitments" },
  { text: "Remains constructive under pressure" },
  { text: "Sets direction and priorities" },
  { text: "Facilitates collective momentum across teams" },
  { text: "Removes obstacles to progress" },
  { text: "Summarizes what was said to confirm understanding" },
  { text: "Develops others for long-term impact" },
  { text: "Connects work to vision and strategy" },
  { text: "Gives others ownership within boundaries" },
  { text: "Creates environment for others to thrive and perform" },
  { text: "Stops work that no longer adds value" },
  { text: "Engages multiple perspectives" },
  { text: "Seeks input from affected stakeholders" },
  { text: "Sets high standards for performance" },
  { text: "Delegates effectively to empower others" },
  { text: "Promotes a culture of inclusion and respect" },
  { text: "Anticipates future trends and challenges" },
] as const;

async function main() {
  console.log("🧹 Cleaning all existing responses and tokens...");
  
  // Delete in order to avoid FK issues
  await prisma.answer.deleteMany();
  await prisma.response.deleteMany();
  await prisma.inviteToken.deleteMany();
  await prisma.option.deleteMany();
  await prisma.question.deleteMany();
  await prisma.form.deleteMany();
  await prisma.participant.deleteMany();

  console.log("👤 Creating Test Participants...");
  
  const alice = await prisma.participant.create({
    data: {
      fullName: "Alice Andersen",
      email: "alice@example.com",
      employeeId: "EMP-001",
      department: "Engineering",
      designation: "Senior Engineer",
    },
  });

  const bob = await prisma.participant.create({
    data: {
      fullName: "Bob Berglund",
      email: "bob@example.com",
      employeeId: "EMP-002",
      department: "Product",
      designation: "Product Manager",
    },
  });

  console.log("📋 Creating Form and Questions...");
  
  const form = await prisma.form.create({
    data: { title: "360° Leadership Behaviour Survey" },
  });

  for (const q of QUESTIONS) {
    const question = await prisma.question.create({
      data: { text: q.text, formId: form.id },
    });

    await prisma.option.createMany({
      data: OPTIONS.map((opt) => ({
        label: opt.label,
        value: opt.value ?? null,
        questionId: question.id,
      })),
    });
  }

  console.log("\n✨ DATABASE IS READY FOR MANUAL TESTING!");
  console.log("------------------------------------------");
  console.log(`Form ID: ${form.id}`);
  console.log(`Alice (Participant): alice@example.com`);
  console.log(`Bob   (Participant): bob@example.com`);
  console.log("\nDIRECTIONS:");
  console.log("1. Go to /dashboard/send-invite");
  console.log("2. Select Alice as participant.");
  console.log("3. Send an invite to your own email.");
  console.log("4. Send another invite to a different email.");
  console.log("5. Fill both, then generate the report! (The system automatically identifies self-responses via email matching).");
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });