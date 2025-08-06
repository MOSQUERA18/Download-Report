import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import time
import os

class SenaAutomation:
    def __init__(self):
        self.setup_driver()
            
    def setup_driver(self):
        """Configura el driver de Chrome con opciones para manejar SSL"""
        chrome_options = Options()
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--allow-running-insecure-content')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--ignore-certificate-errors-spki-list')
        
        # Configurar directorio de descarga
        download_dir = os.path.join(os.getcwd(), "reportes_sena")
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
            
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 20)

    def read_excel_fichas(self, excel_path):
        """Lee el archivo Excel y extrae los números de ficha"""
        try:
            df = pd.read_excel(excel_path)
            fichas = df.iloc[:, 0].tolist()
            print(f"Se encontraron {len(fichas)} fichas en el Excel")
            return fichas
        except Exception as e:
            print(f"Error al leer el Excel: {e}")
            return []

    def switch_to_login_iframe(self):
        """Cambia al iframe donde está el formulario de login"""
        try:
            print("Buscando iframe de login...")
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
            
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            print(f"Se encontraron {len(iframes)} iframes")
            
            for i, iframe in enumerate(iframes):
                try:
                    print(f"Probando iframe {i+1}...")
                    self.driver.switch_to.frame(iframe)
                    
                    try:
                        self.driver.find_element(By.ID, "username")
                        print(f"¡Formulario de login encontrado en iframe {i+1}!")
                        return True
                    except:
                        self.driver.switch_to.default_content()
                        continue
                        
                except Exception as e:
                    print(f"Error al acceder al iframe {i+1}: {e}")
                    self.driver.switch_to.default_content()
                    continue
            
            print("No se encontró el iframe con el formulario de login")
            return False
            
        except Exception as e:
            print(f"Error al buscar iframes: {e}")
            return False

    def navigate_to_sena(self):
        """Navega al sitio de SENA Sofia Plus y hace login"""
        try:
            print("Navegando a SENA Sofia Plus...")
            self.driver.get("http://senasofiaplus.edu.co/sofia-public/")
            time.sleep(3)
            
            try:
                continue_button = self.driver.find_element(By.ID, "proceed-button")
                continue_button.click()
                print("Advertencia SSL manejada")
                time.sleep(2)
            except:
                print("No se encontró advertencia SSL o ya se pasó")
            
            if not self.switch_to_login_iframe():
                return False
            
            if not self.login():
                return False
            
            if not self.post_login_navigation():
                return False
            
            return True
            
        except Exception as e:
            print(f"Error al navegar al sitio: {e}")
            return False

    def login(self):
        """Realiza el proceso de login"""
        try:
            print("Iniciando proceso de login...")
            self.wait.until(EC.presence_of_element_located((By.ID, "username")))
            
            documento_field = self.driver.find_element(By.ID, "username")
            documento_field.clear()
            documento_field.send_keys("38363049")
            print("Número de documento ingresado")
            
            password_field = self.driver.find_element(By.NAME, "josso_password")
            password_field.clear()
            password_field.send_keys("10337684990Ad*")
            print("Contraseña ingresada")
            
            ingresar_button = self.driver.find_element(By.CSS_SELECTOR, "input.login100-form-btn[value='Ingresar']")
            ingresar_button.click()
            print("Botón INGRESAR presionado")
            
            time.sleep(5)
            self.driver.switch_to.default_content()
            print("Volviendo al contexto principal")
            
            try:
                time.sleep(3)
                current_url = self.driver.current_url
                print(f"URL actual después del login: {current_url}")
                
                if "login" not in current_url.lower() or len(self.driver.find_elements(By.TAG_NAME, "iframe")) > 0:
                    print("Login exitoso")
                    return True
                else:
                    print("Login falló - aún en página de login")
                    return False
                    
            except Exception as e:
                print(f"Error verificando login: {e}")
                return False
                
        except Exception as e:
            print(f"Error durante el login: {e}")
            self.driver.switch_to.default_content()
            return False

    def post_login_navigation(self):
        """Realiza la navegación después del login"""
        try:
            if not self.select_role():
                return False
            
            if not self.navigate_to_inscripcion():
                return False
            
            return True
        except Exception as e:
            print(f"Error en la navegación post-login: {e}")
            return False

    def select_role(self):
        """Selecciona el rol después del login"""
        try:
            print("Seleccionando rol...")
            role_select = self.wait.until(
                EC.presence_of_element_located((By.ID, "seleccionRol:roles"))
            )
            
            select = Select(role_select)
            select.select_by_value("33")
            print("Rol 'Encargado de ingreso centro formación' seleccionado")
            time.sleep(3)
            return True
            
        except Exception as e:
            print(f"Error al seleccionar rol: {e}")
            return False

    def navigate_to_inscripcion(self):
        """Navega a la opción 'Inscripción' en el menú principal"""
        try:
            print("Navegando a 'Inscripción'...")
            
            inscripcion_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[@class='menuPrimario' and text()='Inscripción']"))
            )
            inscripcion_link.click()
            print("Clic en 'Inscripción' exitoso")
            time.sleep(3)
            
            consultas_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Consultas')]"))
            )
            consultas_link.click()
            print("Clic en 'Consultas' exitoso")
            time.sleep(2)
            
            generar_reporte_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Generar Reporte de Inscripción')]"))
            )
            generar_reporte_link.click()
            print("Clic en 'Generar Reporte de Inscripción' exitoso")
            time.sleep(3)
            
            return True
            
        except Exception as e:
            print(f"Error al navegar a 'Inscripción': {e}")
            return False

    def click_consultar_ficha_button(self):
        """Encuentra y hace clic en el botón 'Consultar ficha'"""
        try:
            print("Buscando el botón 'Consultar ficha'...")
            self.driver.switch_to.default_content()
            time.sleep(2)
            
            outer_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            
            for i, outer in enumerate(outer_iframes):
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame(outer)
                
                if self.try_click_ficha_button_in_current_frame():
                    print(f"✓ Clic exitoso en 'Consultar ficha' en iframe {i+1}")
                    return True
                
                inner_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                for j, inner in enumerate(inner_iframes):
                    try:
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame(outer)
                        self.driver.switch_to.frame(inner)
                        
                        if self.try_click_ficha_button_in_current_frame():
                            print(f"✓ Clic exitoso en 'Consultar ficha' en iframe {i+1}.{j+1}")
                            return True
                            
                    except Exception as inner_e:
                        continue
            
            print("❌ No se encontró el botón 'Consultar ficha'")
            return False
            
        except Exception as e:
            print(f"❌ Error al buscar el botón: {e}")
            return False

    def try_click_ficha_button_in_current_frame(self):
        """Intenta encontrar y hacer clic en el botón 'Consultar ficha'"""
        try:
            selectors = [
                "//img[@title='Consultar ficha']",
                "//input[@title='Consultar ficha']",
                "//button[contains(text(), 'Consultar ficha')]",
                "//a[contains(text(), 'Consultar ficha')]",
                "//*[contains(@title, 'Consultar ficha')]"
            ]
            
            for selector in selectors:
                try:
                    boton = self.driver.find_element(By.XPATH, selector)
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                    time.sleep(0.5)
                    boton.click()
                    print(f"✓ Clic exitoso con selector: {selector}")
                    return True
                except:
                    continue
                    
            return False
            
        except:
            return False

    def wait_for_form_and_insert_ficha(self, ficha):
        """Espera a que aparezca el formulario y procesa la ficha - SIN BUSCAR OVERLAY MODAL"""
        try:
            print(f"Esperando formulario para ficha: {ficha}")
            
            # PASO 1: Esperar un poco para que el formulario se cargue después del clic
            print("Esperando que se cargue el formulario...")
            time.sleep(3)
            
            # PASO 2: Buscar el campo de input directamente en el contexto actual
            print("Buscando el campo de input en el contexto actual...")
            input_ficha = None
            
            try:
                # Intentar encontrar el input en el iframe actual (donde se hizo el clic)
                input_ficha = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "form:codigoFichaITX"))
                )
                print("✓ Campo de input encontrado en el iframe actual")
                
                # Procesar el input
                return self.process_input_field(input_ficha, ficha)
                
            except:
                print("❌ No se encontró el input en el iframe actual, buscando en otros contextos...")
                
                # PASO 3: Si no se encuentra en el iframe actual, buscar en todos los contextos
                return self.find_and_process_input_in_all_contexts(ficha)
                
        except Exception as e:
            print(f"❌ Error general al procesar ficha {ficha}: {e}")
            try:
                self.driver.save_screenshot(f"error_ficha_{ficha}_form.png")
                print(f"Screenshot guardado: error_ficha_{ficha}_form.png")
            except:
                pass
            return False

    def find_and_process_input_in_all_contexts(self, ficha):
        """Busca el campo de input en todos los contextos posibles"""
        try:
            print("Buscando el campo de input en todos los contextos...")
            
            # Contexto 1: Página principal
            print("Probando contexto principal...")
            self.driver.switch_to.default_content()
            try:
                input_ficha = self.driver.find_element(By.ID, "form:codigoFichaITX")
                if input_ficha.is_displayed():
                    print("✓ Campo encontrado en contexto principal")
                    return self.process_input_field(input_ficha, ficha)
            except:
                pass
            
            # Contexto 2: Buscar en todos los iframes
            print("Buscando en todos los iframes...")
            self.driver.switch_to.default_content()
            all_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            
            for i, iframe in enumerate(all_iframes):
                try:
                    iframe_id = iframe.get_attribute('id')
                    print(f"Probando iframe {i+1} (ID: {iframe_id})...")
                    
                    # Cambiar al iframe
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame(iframe)
                    
                    # Buscar el input
                    try:
                        input_ficha = self.driver.find_element(By.ID, "form:codigoFichaITX")
                        if input_ficha.is_displayed():
                            print(f"✓ Campo encontrado en iframe {i+1}")
                            return self.process_input_field(input_ficha, ficha)
                    except:
                        pass
                    
                    # Buscar en iframes anidados
                    nested_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                    for j, nested in enumerate(nested_iframes):
                        try:
                            print(f"  Probando iframe anidado {i+1}.{j+1}...")
                            self.driver.switch_to.default_content()
                            self.driver.switch_to.frame(iframe)
                            self.driver.switch_to.frame(nested)
                            
                            input_ficha = self.driver.find_element(By.ID, "form:codigoFichaITX")
                            if input_ficha.is_displayed():
                                print(f"✓ Campo encontrado en iframe anidado {i+1}.{j+1}")
                                return self.process_input_field(input_ficha, ficha)
                                
                        except:
                            continue
                        
                except Exception as e:
                    continue
            
            print("❌ No se encontró el campo de input en ningún contexto")
            self.driver.switch_to.default_content()
            return False
            
        except Exception as e:
            print(f"Error buscando en contextos: {e}")
            self.driver.switch_to.default_content()
            return False

    def process_input_field(self, input_element, ficha):
        """Procesa el campo de input insertando la ficha"""
        try:
            print(f"Procesando campo de input para ficha: {ficha}")
            
            # Verificar que el elemento sea visible e interactuable
            if not input_element.is_displayed():
                print("❌ El campo de input no es visible")
                return False
            
            if not input_element.is_enabled():
                print("❌ El campo de input no está habilitado")
                return False
            
            # Hacer scroll al elemento
            self.driver.execute_script("arguments[0].scrollIntoView(true);", input_element)
            time.sleep(0.5)
            
            # Hacer clic para enfocar el campo
            print("Haciendo clic en el campo de input...")
            input_element.click()
            time.sleep(0.5)
            
            # Limpiar el campo
            print("Limpiando el campo...")
            input_element.clear()
            time.sleep(0.5)
            
            # Insertar la ficha
            print(f"Insertando ficha: {ficha}")
            input_element.send_keys(str(ficha))
            time.sleep(1)
            
            # Verificar que se insertó correctamente
            inserted_value = input_element.get_attribute('value')
            if inserted_value == str(ficha):
                print(f"✓ Ficha '{ficha}' insertada correctamente")
                time.sleep(2)
                return True
            else:
                print(f"❌ Error: Se esperaba '{ficha}' pero se insertó '{inserted_value}'")
                return False
                
        except Exception as e:
            print(f"❌ Error al procesar el campo de input: {e}")
            return False

    def process_single_ficha(self, ficha):
        """Procesa una sola ficha: hace clic en el botón y luego inserta la ficha"""
        try:
            print(f"\n{'='*50}")
            print(f"PROCESANDO FICHA: {ficha}")
            print(f"{'='*50}")
            
            # Paso 1: Hacer clic en el botón "Consultar ficha"
            if not self.click_consultar_ficha_button():
                print(f"❌ No se pudo hacer clic en 'Consultar ficha' para la ficha {ficha}")
                return False
            
            # Paso 2: Esperar el formulario e insertar la ficha (SIN BUSCAR OVERLAY)
            if not self.wait_for_form_and_insert_ficha(ficha):
                print(f"❌ No se pudo insertar la ficha {ficha}")
                return False
            
            print(f"✓ Ficha {ficha} procesada exitosamente")
            return True
            
        except Exception as e:
            print(f"❌ Error al procesar ficha {ficha}: {e}")
            return False

    def run_automation(self, excel_path):
        """Ejecuta la automatización completa"""
        try:
            # Leer fichas del Excel
            fichas = self.read_excel_fichas(excel_path)
            if not fichas:
                return
            
            # Tomar solo las primeras 3 fichas para testing
            fichas = fichas[:3]  # Comentar esta línea para procesar todas
            
            # Navegar al sitio (una sola vez)
            if not self.navigate_to_sena():
                print("❌ Error en la navegación inicial")
                return
            
            # Procesar cada ficha
            successful = 0
            failed = 0
            
            for i, ficha in enumerate(fichas):
                print(f"\n--- Procesando ficha {i+1} de {len(fichas)} ---")
                
                if self.process_single_ficha(ficha):
                    successful += 1
                else:
                    failed += 1
                
                # Pausa entre fichas
                if i < len(fichas) - 1:  # No pausar después de la última
                    print("Esperando antes de la siguiente ficha...")
                    time.sleep(3)
            
            print(f"\n{'='*50}")
            print(f"AUTOMATIZACIÓN COMPLETADA")
            print(f"Exitosas: {successful}")
            print(f"Fallidas: {failed}")
            print(f"{'='*50}")
            
        except Exception as e:
            print(f"Error en la automatización: {e}")
        finally:
            input("Presiona Enter para cerrar el navegador...")
            self.driver.quit()

def main():
    excel_path = "fichas_sena.xlsx"
    automation = SenaAutomation()
    automation.run_automation(excel_path)

if __name__ == "__main__":
    main()
