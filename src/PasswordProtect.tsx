import { useState } from 'react'

export default function PasswordProtect({ children }: { children: React.ReactNode }) {
  const [password, setPassword] = useState('')
  const [isAuthenticated, setIsAuthenticated] = useState(
    sessionStorage.getItem('authenticated') === 'true'
  )

  const correctPassword = 'AlexBonnette'

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (password === correctPassword) {
      setIsAuthenticated(true)
      sessionStorage.setItem('authenticated', 'true')
    } else {
      alert('Incorrect password')
    }
  }

  if (isAuthenticated) {
    return <>{children}</>
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 flex items-center justify-center p-4">
      <div className="bg-slate-800 rounded-2xl p-8 border border-slate-700 max-w-md w-full">
        <h1 className="text-3xl font-bold text-white mb-6 text-center">
          Facility Calculator
        </h1>
        <p className="text-slate-300 mb-6 text-center">
          Enter password to access
        </p>
        <form onSubmit={handleSubmit}>
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
