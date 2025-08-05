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
        self.wait = WebDriverWait(self.driver, 20) # Aumentado a 20 segundos

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
                    print(f"Clic exitoso en 'Consultar ficha' en iframe {i+1}")
                    return True
                
                inner_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                for j, inner in enumerate(inner_iframes):
                    try:
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame(outer)
                        self.driver.switch_to.frame(inner)
                        
                        if self.try_click_ficha_button_in_current_frame():
                            print(f"Clic exitoso en 'Consultar ficha' en iframe {i+1}.{j+1}")
                            return True
                            
                    except Exception as inner_e:
                        continue
            
            print("No se encontró el botón 'Consultar ficha'")
            return False
            
        except Exception as e:
            print(f"Error al buscar el botón: {e}")
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
                    print(f"Clic exitoso con selector: {selector}")
                    return True
                except:
                    continue
                    
            return False
            
        except:
            return False

    def capture_full_html_debug(self):
        """Captura TODO el HTML de la página para análisis"""
        try:
            print("\n=== CAPTURANDO HTML COMPLETO ===")
            self.driver.switch_to.default_content()
            
            # Capturar HTML completo
            full_html = self.driver.page_source
            
            # Guardar en archivo para análisis
            with open("debug_full_page.html", "w", encoding="utf-8") as f:
                f.write(full_html)
            print("✓ HTML completo guardado en: debug_full_page.html")
            
            # Buscar TODOS los divs con cualquier ID
            all_divs = self.driver.find_elements(By.TAG_NAME, "div")
            print(f"\n=== TODOS LOS DIVS ENCONTRADOS: {len(all_divs)} ===")
            
            modal_candidates = []
            for i, div in enumerate(all_divs):
                try:
                    div_id = div.get_attribute('id')
                    div_class = div.get_attribute('class')
                    div_style = div.get_attribute('style')
                    is_visible = div.is_displayed()
                    
                    # Buscar divs que podrían ser modales
                    if (div_id and ('modal' in div_id.lower() or 'dialog' in div_id.lower() or 'jsp' in div_id.lower())) or \
                       (div_class and ('modal' in div_class.lower() or 'dialog' in div_class.lower())):
                        modal_candidates.append({
                            'index': i,
                            'id': div_id,
                            'class': div_class,
                            'visible': is_visible,
                            'style': div_style[:100] if div_style else 'None'
                        })
                        
                except Exception as e:
                    continue
            
            print(f"\n=== CANDIDATOS A MODAL: {len(modal_candidates)} ===")
            for candidate in modal_candidates:
                print(f"Div {candidate['index']}:")
                print(f"  - ID: {candidate['id']}")
                print(f"  - Class: {candidate['class']}")
                print(f"  - Visible: {candidate['visible']}")
                print(f"  - Style: {candidate['style']}")
                print()
            
            # Buscar TODOS los iframes
            all_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            print(f"\n=== TODOS LOS IFRAMES: {len(all_iframes)} ===")
            
            for i, iframe in enumerate(all_iframes):
                try:
                    iframe_id = iframe.get_attribute('id')
                    iframe_src = iframe.get_attribute('src')
                    iframe_class = iframe.get_attribute('class')
                    is_visible = iframe.is_displayed()
                    
                    print(f"Iframe {i+1}:")
                    print(f"  - ID: {iframe_id}")
                    print(f"  - Src: {iframe_src[:100] if iframe_src else 'None'}")
                    print(f"  - Class: {iframe_class}")
                    print(f"  - Visible: {is_visible}")
                    
                    # Si el iframe contiene 'ficha' o 'modal', intentar explorarlo
                    if iframe_src and ('ficha' in iframe_src.lower() or 'modal' in iframe_src.lower()):
                        print(f"  *** IFRAME SOSPECHOSO DE SER EL MODAL ***")
                        try:
                            # Guardar contexto actual
                            self.driver.switch_to.frame(iframe)
                            iframe_html = self.driver.page_source
                            
                            # Guardar HTML del iframe
                            with open(f"debug_iframe_{i+1}.html", "w", encoding="utf-8") as f:
                                f.write(iframe_html)
                            print(f"  - HTML del iframe guardado en: debug_iframe_{i+1}.html")
                            
                            # Buscar inputs en este iframe
                            inputs_in_iframe = self.driver.find_elements(By.TAG_NAME, "input")
                            print(f"  - Inputs encontrados en iframe: {len(inputs_in_iframe)}")
                            
                            for j, inp in enumerate(inputs_in_iframe):
                                inp_id = inp.get_attribute('id')
                                inp_name = inp.get_attribute('name')
                                inp_type = inp.get_attribute('type')
                                print(f"    Input {j+1}: id='{inp_id}' name='{inp_name}' type='{inp_type}'")
                            
                            # Volver al contexto principal
                            self.driver.switch_to.default_content()
                            
                        except Exception as iframe_error:
                            print(f"  - Error explorando iframe: {iframe_error}")
                            self.driver.switch_to.default_content()
                    
                    print()
                    
                except Exception as e:
                    print(f"Iframe {i+1}: Error obteniendo atributos - {e}")
            
            print("=== FIN CAPTURA HTML ===\n")
            
        except Exception as e:
            print(f"Error en captura HTML: {e}")

    def process_single_ficha_with_full_debug(self, ficha):
        """Versión con debug completo y lógica de modal mejorada para iframes anidados"""
        try:
            print(f"\n{'='*50}")
            print(f"PROCESANDO FICHA: {ficha}")
            print(f"{'='*50}")
            
            # Paso 1: Hacer clic en el botón "Consultar ficha"
            if not self.click_consultar_ficha_button():
                print(f"❌ No se pudo hacer clic en 'Consultar ficha' para la ficha {ficha}")
                return False
            
            # Paso 2: Esperar más tiempo para que el modal se abra completamente y capturar HTML
            print("Esperando 5 segundos para que el modal se abra completamente...")
            time.sleep(5)
            self.capture_full_html_debug() # Capturar HTML para análisis
            
            # Paso 3: Intentar encontrar el contenedor principal del modal (myModal o myModal2)
            print("Intentando encontrar el contenedor principal del modal (myModal o myModal2) visible...")
            self.driver.switch_to.default_content() # Asegurarse de estar en el contexto principal
            
            main_modal_container = None
            modal_container_found = False
            
            # Estrategia 1: Buscar por ID 'myModal' o 'myModal2' si son visibles
            modal_ids = ["myModal", "myModal2"]
            for modal_id in modal_ids:
                try:
                    main_modal_container = self.wait.until(
                        EC.visibility_of_element_located((By.ID, modal_id))
                    )
                    print(f"✓ Contenedor principal '{modal_id}' encontrado y visible.")
                    modal_container_found = True
                    break
                except:
                    continue
            
            # Estrategia 2: Buscar por clase 'modal fade' y estilo 'display: block' si no se encontró por ID
            if not modal_container_found:
                try:
                    main_modal_container = self.wait.until(
                        EC.visibility_of_element_located((By.XPATH, "//div[contains(@class, 'modal') and contains(@class, 'fade') and contains(@style, 'display: block')]"))
                    )
                    print("✓ Contenedor principal encontrado por clase 'modal fade' y estilo 'display: block'.")
                    modal_container_found = True
                except:
                    print("❌ No se encontró un contenedor de modal principal visible.")
                    return False
            
            if not modal_container_found:
                print("❌ No se pudo encontrar el contenedor principal del modal.")
                return False

            # Paso 4: Buscar el iframe interno (ifraModal o ifraModal2) dentro del contenedor principal del modal
            print("Buscando el iframe interno (ifraModal o ifraModal2) dentro del modal principal...")
            inner_iframe_element = None
            inner_iframe_found = False
            
            iframe_ids = ["ifraModal", "ifraModal2"]
            for iframe_id in iframe_ids:
                try:
                    inner_iframe_element = main_modal_container.find_element(By.ID, iframe_id)
                    print(f"✓ Iframe interno '{iframe_id}' encontrado dentro del modal principal.")
                    inner_iframe_found = True
                    break
                except:
                    continue
            
            if not inner_iframe_found:
                # Fallback: buscar cualquier iframe dentro del contenedor principal
                try:
                    inner_iframe_element = main_modal_container.find_element(By.TAG_NAME, "iframe")
                    print("✓ Iframe genérico encontrado dentro del modal principal.")
                    inner_iframe_found = True
                except:
                    print("❌ No se encontró ningún iframe interno dentro del modal principal.")
                    return False
            
            if not inner_iframe_found:
                print("❌ No se pudo encontrar el iframe interno del modal.")
                return False

            # Paso 5: Cambiar al iframe interno (ifraModal/ifraModal2)
            print("Cambiando al iframe interno (ifraModal/ifraModal2)...")
            try:
                self.driver.switch_to.frame(inner_iframe_element)
                print("✓ Cambio exitoso al iframe interno.")
            except Exception as e:
                print(f"❌ Error al cambiar al iframe interno: {e}")
                self.driver.switch_to.default_content()
                return False

            # Paso 6: Esperar que el iframe anidado (modalDialogContentviewDialog2) aparezca dentro del iframe interno
            print("Esperando que el iframe anidado 'modalDialogContentviewDialog2' aparezca dentro del iframe interno...")
            nested_iframe_element = None
            try:
                nested_iframe_element = self.wait.until(
                    EC.presence_of_element_located((By.ID, "modalDialogContentviewDialog2"))
                )
                print("✓ Iframe anidado 'modalDialogContentviewDialog2' encontrado.")
            except Exception as e:
                print(f"❌ No se encontró el iframe anidado 'modalDialogContentviewDialog2': {e}")
                self.driver.switch_to.default_content()
                return False

            # Paso 7: Cambiar al iframe anidado (modalDialogContentviewDialog2)
            print("Cambiando al iframe anidado 'modalDialogContentviewDialog2'...")
            try:
                self.driver.switch_to.frame(nested_iframe_element)
                print("✓ Cambio exitoso al iframe anidado.")
            except Exception as e:
                print(f"❌ Error al cambiar al iframe anidado: {e}")
                self.driver.switch_to.default_content()
                return False

            # Paso 8: Localizar y interactuar con el campo de input dentro del iframe anidado
            print("Buscando campo de input 'form:codigoFichaITX' dentro del iframe anidado...")
            input_ficha = None
            input_selectors = [
                (By.ID, "form:codigoFichaITX"),
                (By.NAME, "form:codigoFichaITX"),
                (By.XPATH, "//input[@type='text']"),
                (By.XPATH, "//input[contains(@title, 'Ficha')]")
            ]

            for by_type, selector in input_selectors:
                try:
                    input_ficha = self.wait.until(EC.presence_of_element_located((by_type, selector)))
                    print(f"✓ Input encontrado con selector: {selector}")
                    break
                except:
                    continue
            
            if not input_ficha:
                print("❌ No se pudo encontrar el campo de input en el iframe anidado.")
                self.driver.switch_to.default_content()
                return False

            # Paso 9: Procesar el input
            self.driver.execute_script("arguments[0].scrollIntoView(true);", input_ficha)
            time.sleep(0.5)
            input_ficha.click()
            time.sleep(0.5)
            input_ficha.clear()
            input_ficha.send_keys(str(ficha))
            time.sleep(1)

            # Paso 10: Verificar y volver al contexto principal
            inserted_value = input_ficha.get_attribute('value')
            if inserted_value == str(ficha):
                print(f"✓ Ficha '{ficha}' insertada correctamente.")
                time.sleep(2)
                self.driver.switch_to.default_content()
                return True
            else:
                print(f"❌ Error: Se esperaba '{ficha}' pero se insertó '{inserted_value}'.")
                self.driver.switch_to.default_content()
                return False

        except Exception as e:
            print(f"❌ Error general al procesar ficha {ficha}: {e}")
            try:
                self.driver.save_screenshot(f"error_ficha_{ficha}_detailed.png")
                print(f"Screenshot guardado: error_ficha_{ficha}_detailed.png")
            except:
                pass
            self.driver.switch_to.default_content()
            return False

    def process_single_ficha(self, ficha):
        """Procesa una sola ficha - USA EL DEBUG COMPLETO"""
        return self.process_single_ficha_with_full_debug(ficha)

    def run_automation(self, excel_path):
        """Ejecuta la automatización completa"""
        try:
            # Leer fichas del Excel
            fichas = self.read_excel_fichas(excel_path)
            if not fichas:
                return
            
            # Tomar solo las primeras 1 ficha para testing
            fichas = fichas[:1]  # Comentar esta línea para procesar todas
            
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
