import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();

async function main() {

  // ✅ Clear old data (optional)
  await prisma.answer.deleteMany();
  await prisma.response.deleteMany();
  await prisma.option.deleteMany();
  await prisma.question.deleteMany();
  await prisma.form.deleteMany();

  // ✅ Create Form
  const form = await prisma.form.create({
    data: {
      title: "Leadership Behavior Survey",
    },
  });

  // ✅ Standard Options
  const optionsData = [
    { label: "Rarely", value: 25 },
    { label: "Sometimes", value: 50 },
    { label: "Often", value: 75 },
    { label: "Always", value: 100 },
    { label: "Insufficient Exposure", value: null },
  ];

  // ✅ Questions List
  const questionsText = [
    "Takes ownership of outcomes end-to-end",
    "Prioritizes highest-value work",
    "Makes timely decisions",
    "Questions assumptions and ways of working",
    "Adapts based on feedback",
    "Runs continuous experiments instead of waiting",
    "Uses external signals to guide decisions",
    "Proactively seeks feedback to improve impact",
  ];

  // ✅ Create Questions + Options
  for (const text of questionsText) {
    const question = await prisma.question.create({
      data: {
        text,
        formId: form.id,
      },
    });

    await prisma.option.createMany({
      data: optionsData.map((opt) => ({
        label: opt.label,
        value: opt.value,
        questionId: question.id,
      })),
    });
  }

  console.log("✅ Seeding completed!");
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });