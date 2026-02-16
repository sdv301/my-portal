// src/App.jsx
import { useState, useCallback, useEffect } from 'react';
import Map from './components/Map';
import Filters from './components/Filters';
import InfoPanel from './components/InfoPanel';
import './App.css';

function App() {
  const [filters, setFilters] = useState({
    startDate: '',
    endDate: '',
    dataType: 'all'
  });
  const [selectedDistrict, setSelectedDistrict] = useState(null);
  const [isFiltersExpanded, setIsFiltersExpanded] = useState(false);

  // ДОБАВЛЕНО: Автоматически тянем самую старую и новую дату из базы при запуске
  useEffect(() => {
    fetch('http://localhost:5000/api/date-range')
      .then(res => res.json())
      .then(data => {
        if (data.min_year && data.max_year) {
          setFilters(prev => ({
            ...prev,
            startDate: `${data.min_year}-01-01`,
            endDate: `${data.max_year}-12-31`
          }));
        }
      })
      .catch(err => console.error('Ошибка загрузки дат:', err));
  }, []);

  const handleFilterChange = useCallback((newFilters) => {
    setFilters(prev => ({ ...prev, ...newFilters }));
  }, []);

  const handleDistrictSelect = useCallback((district) => {
    const nameToId = {
      'город Якутск': 'yakutsk',
      'Жатай': 'zhatay',
      'Абыйский район': 'abysky',
      'Алданский район': 'aldansky',
      'Аллаиховский район': 'allaikhovsky',
      'Амгинский район': 'amginsky',
      'Анабарский район': 'anabarsky',
      'Булунский район': 'bulunsky',
      'Верхневилюйский район': 'verkhnevilyuysky',
      'Верхнеколымский район': 'verkhnekolymsky',
      'Верхоянский район': 'verkhoyansky',
      'Вилюйский район': 'vilyuysky',
      'Горный район': 'gorny',
      'Жиганский район': 'zhigansky',
      'Кобяйский район': 'kobyaysky',
      'Ленский район': 'lensky',
      'Мегино-Кангаласский район': 'megino-kangalassky',
      'Мирнинский район': 'mirninsky',
      'Момский район': 'omsky',
      'Намский район': 'namsky',
      'Нерюнгринский район': 'neryungrinsky',
      'Нижнеколымский район': 'nizhnekolymsky',
      'Нюрбинский район': 'nyurbinsky',
      'Оймяконский район': 'oymyakonsky',
      'Олёкминский район': 'olekminsky',
      'Оленёкский район': 'olenek',
      'Среднеколымский район': 'srednekolymsky',
      'Сунтарский район': 'suntarsky',
      'Таттинский район': 'tattinsky',
      'Томпонский район': 'tomponsky',
      'Усть-Алданский район': 'ust-aldansky',
      'Усть-Майский район': 'ust-maysky',
      'Усть-Янский район': 'ust-yansky',
      'Хангаласский район': 'khangalassky',
      'Чурапчинский район': 'churapchinsky',
      'Эвено-Бытантайский район': 'eveno-bytantaysky'
    };

    const id = nameToId[district.name] || district.id;
    setSelectedDistrict({ id, name: district.name });
  }, []);

  const handleFiltersToggle = useCallback((expanded) => {
    setIsFiltersExpanded(expanded);
  }, []);

  const handleInfoPanelToggle = useCallback((expanded) => {
    if (expanded) setIsFiltersExpanded(false);
  }, []);

  return (
    <div className="app">
      <Filters 
        filters={filters} 
        setFilters={handleFilterChange}
        selectedDistrict={selectedDistrict}
        onPanelToggle={handleFiltersToggle}
      />
      
      <div className="main">
        <div className="map-container">
          <Map filters={filters} onDistrictClick={handleDistrictSelect} />
        </div>
        <InfoPanel 
          district={selectedDistrict} 
          filters={filters}
          isFiltersExpanded={isFiltersExpanded}
          onPanelToggle={handleInfoPanelToggle}
        />
      </div>

      {/* Кнопка админки */}
      <a href="/admin.html" className="hidden-admin-btn" title="Панель управления">⚙️</a>
    </div>
  );
}

export default App;