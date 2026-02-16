// src/components/Filters.jsx
import React, { useState, useEffect } from 'react';
import './Filters.css';

const Filters = ({ filters, setFilters, selectedDistrict, onPanelToggle }) => {
  const [isExpanded, setIsExpanded] = useState(false);
  const [districts, setDistricts] = useState([]);
  
  // Отдельные состояния для панели экспорта Excel
  const [exportDistrict, setExportDistrict] = useState('all');
  const [exportStart, setExportStart] = useState('');
  const [exportEnd, setExportEnd] = useState('');
  const [alertMsg, setAlertMsg] = useState(null);

  useEffect(() => {
    fetch('http://localhost:5000/api/districts')
      .then(res => res.json())
      .then(data => setDistricts(data))
      .catch(err => console.error(err));
  }, []);

  useEffect(() => {
    if (selectedDistrict?.id) {
      setExportDistrict(selectedDistrict.id);
    }
  }, [selectedDistrict]);

  const togglePanel = (e) => {
    if (e) e.stopPropagation();
    const newState = !isExpanded;
    setIsExpanded(newState);
    if (onPanelToggle) onPanelToggle(newState);
  };

  // Умная логика выгрузки с проверкой наличия данных
  const handleExport = async () => {
    setAlertMsg(null); // Сбрасываем старое сообщение
    
    try {
      const res = await fetch(`http://localhost:5000/api/check-export?startDate=${exportStart}&endDate=${exportEnd}&district_id=${exportDistrict}`);
      const data = await res.json();
      
      if (data.hasData) {
        // Данные есть - начинаем скачивание
        window.location.href = `http://localhost:5000/api/export-excel?startDate=${exportStart}&endDate=${exportEnd}&district_id=${exportDistrict}`;
      } else {
        // Данных нет - показываем алерт и подставляем последние доступные
        const latest = data.latest_year;
        if (latest) {
           setAlertMsg(`⚠️ В выбранный период выдачи МЦ не было. Самые свежие данные найдены за ${latest} год.`);
           setExportStart(`${latest}-01-01`);
           setExportEnd(`${latest}-12-31`);
        } else {
           setAlertMsg('❌ Для этого района в базе данных вообще нет записей о выдаче имущества.');
        }
      }
    } catch (err) {
      console.error(err);
      setAlertMsg('❌ Ошибка связи с сервером.');
    }
  };

  return (
    <div className="filters-wrapper">
      <button className="panel-toggle" onClick={togglePanel}>
        <span>{isExpanded ? '✕ Скрыть панель' : '⚙️ Настройки и Выгрузка'}</span>
      </button>

      <div className={`filters-panel ${isExpanded ? 'expanded' : ''}`}>
        <div className="filters-grid">
          
          {/* Левая колонка: Настройки отображения карты (как было) */}
          <div className="filter-section">
            <h3>🗺️ Отображение на карте</h3>
            <p style={{fontSize: '13px', color: 'var(--text-muted)', marginBottom: '15px'}}>
              Эти даты фильтруют информацию, которая появляется в боковой панели при клике на районы.
            </p>
            <div className="date-inputs">
              <div className="filter-group">
                <label>Начало периода</label>
                <input 
                  type="date" 
                  value={filters.startDate || ''} 
                  onChange={e => setFilters({...filters, startDate: e.target.value})} 
                />
              </div>
              <div className="filter-group">
                <label>Конец периода</label>
                <input 
                  type="date" 
                  value={filters.endDate || ''} 
                  onChange={e => setFilters({...filters, endDate: e.target.value})} 
                />
              </div>
            </div>
          </div>
          
          {/* Правая колонка: Экспорт Excel (Обновленная) */}
          <div className="export-section">
            <h3>📊 Выгрузка отчета (Excel)</h3>
            
            <div className="filter-group" style={{marginBottom: '10px'}}>
              <label>Район для выгрузки:</label>
              <select value={exportDistrict} onChange={e => setExportDistrict(e.target.value)}>
                <option value="all">📁 Все районы (Полный свод)</option>
                {districts.map(d => (
                  <option key={d.id} value={d.id}>📍 {d.name}</option>
                ))}
              </select>
            </div>

            <div className="date-inputs" style={{marginBottom: '10px'}}>
              <div className="filter-group">
                <label>Выгрузить с (дата)</label>
                <input type="date" value={exportStart} onChange={e => setExportStart(e.target.value)} />
              </div>
              <div className="filter-group">
                <label>Выгрузить по (дата)</label>
                <input type="date" value={exportEnd} onChange={e => setExportEnd(e.target.value)} />
              </div>
            </div>

            {alertMsg && (
              <div className="alert-box">
                {alertMsg}
              </div>
            )}

            <button className="export-btn" onClick={handleExport}>
              📥 Скачать ведомость
            </button>
          </div>

        </div>
      </div>
    </div>
  );
};

export default Filters;