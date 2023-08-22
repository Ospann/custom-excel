import Excel from './components/Excel';

const App = () => {
  const data = [
    { "№": "1", "Наименование": "W-088-T Lacoste Pour Femme By Zeyd (масло)", "Кол-во": 30, "Ед.": "", "Цена": 92.00, "Сумма": 18400.00 },
    { "№": "2", "Наименование": "M-011-T Chanel Bleu De Chanel (масло)", "Кол-во": 30, "Ед.": "", "Цена": 92.00, "Сумма": 18400.00 },
    // ... more data
  ];

  return (
    <div>
      <h1>Styled Excel Download</h1>
      <Excel data={data} />
    </div>
  );
};

export default App;
