#!/usr/bin/env python3
import requests
import time
import threading
import subprocess
import sys

def create_tunnel():
    """Create a public tunnel to the local Streamlit app"""
    
    print("🚀 Creating public tunnel for Streamlit app...")
    
    # Method 1: Try using pyngrok
    try:
        from pyngrok import ngrok
        print("✅ Using ngrok for tunnel...")
        
        # Start ngrok tunnel
        tunnel = ngrok.connect(8500, "http")
        public_url = tunnel.public_url
        
        print(f"\n🎯 PUBLIEKE URL BESCHIKBAAR:")
        print(f"   {public_url}")
        print(f"\n📱 STREAMLIT APP VERSIE v2.1.0")
        print(f"   - Ultra-flexibele kolom matching")
        print(f"   - Automatische kolom detectie") 
        print(f"   - Openpyxl bug fixes")
        print(f"   - 4 matching strategieën")
        print(f"\n⏰ Tunnel blijft actief...")
        
        # Keep tunnel alive
        try:
            while True:
                time.sleep(30)
                # Check if tunnel is still active
                response = requests.get(f"{public_url}", timeout=5)
                if response.status_code == 200:
                    print(f"✅ {time.strftime('%H:%M:%S')} - Tunnel actief: {public_url}")
                else:
                    print(f"⚠️ {time.strftime('%H:%M:%S')} - Tunnel response: {response.status_code}")
        except KeyboardInterrupt:
            print("\n🛑 Tunnel gestopt door gebruiker")
        except Exception as e:
            print(f"\n❌ Tunnel error: {e}")
        finally:
            ngrok.disconnect(tunnel.public_url)
            ngrok.kill()
        
        return public_url
        
    except ImportError:
        print("❌ pyngrok niet beschikbaar, installeren...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "pyngrok"], check=True)
            print("✅ pyngrok geïnstalleerd, opnieuw proberen...")
            return create_tunnel()  # Retry
        except Exception as e:
            print(f"❌ Kon pyngrok niet installeren: {e}")
    
    # Method 2: Try alternative tunnel methods
    print("🔄 Proberen alternatieve methoden...")
    
    # Method 2a: Try localhost.run via curl
    try:
        print("📡 Proberen localhost.run...")
        result = subprocess.run([
            "curl", "-s", "--max-time", "10",
            f"https://localhost.run/?port=8500"
        ], capture_output=True, text=True, timeout=15)
        
        if result.returncode == 0 and "https://" in result.stdout:
            url = result.stdout.strip()
            print(f"\n🎯 LOCALHOST.RUN URL:")
            print(f"   {url}")
            return url
    except Exception as e:
        print(f"❌ localhost.run gefaald: {e}")
    
    # Method 3: Instructions for manual setup
    print("\n📋 HANDMATIGE TUNNEL SETUP:")
    print("Als je geen publieke URL kunt maken, gebruik dan:")
    print("\n1. **Ngrok (aanbevolen):**")
    print("   - Download ngrok van https://ngrok.com/")
    print("   - Run: ngrok http 8500")
    print("   - Gebruik de gegeven https URL")
    
    print("\n2. **SSH Tunnel:**")
    print("   - ssh -R 80:localhost:8500 serveo.net")
    print("   - Gebruik de gegeven URL")
    
    print("\n3. **Cloudflare Tunnel:**")
    print("   - cloudflared tunnel --url http://localhost:8500")
    
    print(f"\n🔗 LOKALE URL (als je op dezelfde machine zit):")
    print(f"   http://localhost:8500")
    print(f"   http://172.17.0.2:8500")
    
    return None

def check_app_status():
    """Check if the Streamlit app is running"""
    try:
        response = requests.get("http://localhost:8500", timeout=5)
        if response.status_code == 200:
            print("✅ Streamlit app draait op poort 8500")
            return True
        else:
            print(f"⚠️ App response code: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ App niet bereikbaar: {e}")
        return False

def main():
    print("🔬 PUBLIEKE TUNNEL CREATOR voor Excel Taak-Norm Verrijker v2.1.0")
    print("=" * 70)
    
    # Check if app is running
    if not check_app_status():
        print("\n❌ Streamlit app is niet actief op poort 8500")
        print("   Start eerst de app met: streamlit run app_final.py --server.port 8500 --server.address 0.0.0.0")
        return
    
    # Create tunnel
    tunnel_url = create_tunnel()
    
    if tunnel_url:
        print(f"\n🎉 SUCCES! App is publiek beschikbaar via:")
        print(f"   {tunnel_url}")
    else:
        print(f"\n📱 App is lokaal beschikbaar via:")
        print(f"   http://localhost:8500")

if __name__ == "__main__":
    main()