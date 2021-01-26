export type NamedRangeType = {
  comment: string;
  formula: string;
  name: string;
  scope: string;
  type: string;
  value: string;
  visible: boolean;
};

interface NamedRanges {
  items: NamedRangeType[];
}

export const getNamedRanges = async (): Promise<NamedRanges | null> => {
  const result = await Excel.run(async (context) => {
    const { names } = context.workbook;
    names.load();
    await context.sync();

    return JSON.parse(JSON.stringify(names));
  }).catch((error) => {
    console.error(error);
    return null;
  });

  return result;
};
