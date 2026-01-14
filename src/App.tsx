import './index.css'
import { SignedIn, SignedOut, SignIn, UserButton } from '@clerk/clerk-react'
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
      </SignedIn>
    </>
  )
}

export default App