import './index.css'
import FacilityCalculator from './facility-calculator-5'
import PasswordProtect from './PasswordProtect'

function App() {
  return (
    <PasswordProtect>
      <FacilityCalculator />
    </PasswordProtect>
  )
}

export default App