import { getNamedRanges, NamedRangeType } from './common';

const validateNamedRanges = async (): Promise<{
  success: boolean;
  errorCode: string;
  invalidNamedRanges: NamedRangeType[];
}> => {
  const result = await Excel.run(async () => {
    const { items } = await getNamedRanges();
    const invalidNamedRanges = items.filter(({ type }) => type === 'Error');
    return {
      success: true,
      errorCode: '',
      invalidNamedRanges,
    };
  }).catch((error) => {
    console.error(error);
    return {
      success: false,
      errorCode: error.message.split(':')[0],
      invalidNamedRanges: [],
    };
  });

  return result;
};

export default validateNamedRanges;
