import './index.css'
import { SignedIn, SignedOut, SignIn, UserButton } from '@clerk/clerk-react'
import { useState } from 'react'
import FacilityCalculator from './facility-calculator-5'

function App() {
  return (
    <>
      <SignedOut>
        <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 flex items-center justify-center p-4">
          <div className="w-full max-w-md">
            <h1 className="text-3xl font-bold text-white mb-6 text-center">
              Facility Calculator
            </h1>
            <SignIn 
              appearance={{
                elements: {
                  rootBox: "mx-auto",
                  card: "shadow-xl",
                  formButtonPrimary: "bg-blue-600 hover:bg-blue-700",
                  footerActionLink: "text-blue-600 hover:text-blue-700"
                }
              }}
            />
          </div>
        </div>
      </SignedOut>
      
      <SignedIn>
        <PasswordProtectedApp />
      </SignedIn>
    </>
  )
}

function PasswordProtectedApp() {
  const [passwordEntered, setPasswordEntered] = useState(
    sessionStorage.getItem('passwordEntered') === 'true'
  )
  const [password, setPassword] = useState('')

  const correctPassword = 'Fc2026!Tx#9mK$pL7'

  const handlePasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (password === correctPassword) {
      setPasswordEntered(true)
      sessionStorage.setItem('passwordEntered', 'true')
    } else {
      alert('Incorrect password')
      setPassword('')
    }
  }

  if (!passwordEntered) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 flex items-center justify-center p-4">
        <div className="bg-slate-800 rounded-2xl p-8 border border-slate-700 max-w-md w-full">
          <div className="absolute top-4 right-4">
            <UserButton 
              appearance={{
                elements: {
                  avatarBox: "w-10 h-10"
                }
              }}
            />
          </div>
          <h1 className="text-3xl font-bold text-white mb-6 text-center">
            Facility Calculator
          </h1>
          <p className="text-slate-300 mb-6 text-center">
            Enter password to access calculator
          </p>
          <form onSubmit={handlePasswordSubmit}>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Password"
              className="w-full px-4 py-3 bg-slate-700 text-white rounded-lg border border-slate-600 focus:outline-none focus:border-blue-500 mb-4"
              autoFocus
            />
            <button
              type="submit"
              className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-6 rounded-lg"
            >
              Enter
            </button>
          </form>
        </div>
      </div>
    )
  }

  return (
    <>
      <div className="fixed top-4 right-4 z-50">
        <UserButton 
          appearance={{
            elements: {
              avatarBox: "w-10 h-10"
            }
          }}
        />
      </div>
      <FacilityCalculator />
    </>
  )
}

export default App
