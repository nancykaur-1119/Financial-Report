import './App.css';
import IsmReport from './components/IsmReport';

export default function App() {
  return (
    <>
      <header className="app-header">
        <div className="app-header-inner">
          <span className="app-logo">Finance Report</span>
        </div>
      </header>
      <div className="app-container">
        <IsmReport />
      </div>
    </>
  );
}
