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
        


    def seleccionar_primera_opcion(self):
        """Selecciona 'Primera Opción' y devuelve la ruta de iframes donde se encontró"""
        try:
            self.driver.switch_to.default_content()
            time.sleep(1)

            outer_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")

            for i, outer in enumerate(outer_iframes):
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame(outer)

                # Buscar en iframe externo
                try:
                    select_element = self.driver.find_element(By.ID, "opcionesInscritos")
                    Select(select_element).select_by_value("1")
                    print("✓ 'Primera Opción' seleccionada correctamente")
                    return [i]  # Ruta de iframes: solo el externo
                except:
                    # Buscar en iframes internos
                    inner_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                    for j, inner in enumerate(inner_iframes):
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame(outer)
                        self.driver.switch_to.frame(inner)
                        try:
                            select_element = self.driver.find_element(By.ID, "opcionesInscritos")
                            Select(select_element).select_by_value("1")
                            print("✓ 'Primera Opción' seleccionada correctamente (iframe interno)")
                            return [i, j]  # Ruta: externo + interno
                        except:
                            continue

            print("❌ No se encontró el select 'opcionesInscritos'")
            return None

        except Exception as e:
            print(f"❌ Error seleccionando opción: {e}")
            return None
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


    def click_consultar_ficha_button(self):
            """Selecciona 'Primera Opción' y luego hace clic en el botón en el mismo iframe"""
            try:
                ruta_iframes = self.seleccionar_primera_opcion()
                if ruta_iframes is None:
                    print("❌ No se pudo seleccionar la opción.")
                    return False

                # Navegar de nuevo a la misma ruta de iframes
                self.driver.switch_to.default_content()
                for idx, pos in enumerate(ruta_iframes):
                    if idx == 0:
                        outer = self.driver.find_elements(By.TAG_NAME, "iframe")[pos]
                        self.driver.switch_to.frame(outer)
                    elif idx == 1:
                        inner = self.driver.find_elements(By.TAG_NAME, "iframe")[pos]
                        self.driver.switch_to.frame(inner)

                # Buscar el botón en el mismo contexto
                if self.try_click_ficha_button_in_current_frame():
                    print("✓ Clic exitoso en 'Consultar ficha'")
                    return True
                else:
                    print("❌ No se encontró el botón 'Consultar ficha' en el mismo iframe")
                    return False

            except Exception as e:
                print(f"❌ Error al buscar el botón: {e}")
                return False


    def wait_for_form_and_insert_ficha(self, ficha):
        """Espera a que aparezca el formulario y procesa la ficha completa, incluyendo la selección del dropdown en el mismo iframe."""
        try:
            print(f"Esperando formulario para ficha: {ficha}")
            
            # PASO 1: Esperar un poco para que el formulario se cargue después del clic
            print("Esperando que se cargue el formulario...")
            time.sleep(3)
            
            # PASO 2: Buscar e insertar la ficha en el campo de input
            print("Buscando el campo de input en el contexto actual (iframe)...")
            input_ficha = None
            
            try:
                # Intentar encontrar el input en el iframe actual (donde se hizo el clic)
                input_ficha = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "form:codigoFichaITX"))
                )
                print("✓ Campo de input encontrado en el iframe actual")
                
                # Procesar el input
                if not self.process_input_field(input_ficha, ficha):
                    return False
                
            except:
                print("❌ No se encontró el input en el iframe actual, buscando en otros contextos...")
                # Si no se encuentra en el iframe actual, buscar en todos los contextos
                # Esto es un fallback, idealmente debería estar en el iframe correcto
                if not self.find_and_process_input_in_all_contexts(ficha):
                    return False
            
            # PASO 3: Hacer clic en el botón "Consultar" después de insertar la ficha
            print("Buscando el botón 'Consultar' para ejecutar la búsqueda (en iframe)...")
            if not self.click_consultar_button_in_iframe(): # Usar un método específico para el botón dentro del iframe
                print("❌ No se pudo hacer clic en el botón 'Consultar'")
                return False
            
            # PASO 4: Esperar 2 segundos y hacer clic en el botón "Agregar"
            print("Esperando 2 segundos para que aparezcan los resultados (en iframe)...")
            time.sleep(2)
            
            if not self.click_agregar_button():
                print("❌ No se pudo hacer clic en el botón 'Agregar'")
                return False
            
            # PASO 5: Seleccionar "Primera Opción" en el select, que está en el MISMO IFRAME
            print("Buscando el select 'opcionesInscritos' en el mismo iframe...")
            time.sleep(3) # Dar tiempo para que el select aparezca/se actualice
            
            if not self.select_primera_opcion_in_iframe():
                print("❌ No se pudo seleccionar 'Primera Opción' en el iframe")
                return False
            
            print("✓ Proceso completo: ficha insertada, búsqueda ejecutada, elemento agregado y opción seleccionada en el iframe")
            return True
                
        except Exception as e:
            print(f"❌ Error general al procesar ficha {ficha}: {e}")
            try:
                self.driver.save_screenshot(f"error_ficha_{ficha}_form.png")
                print(f"Screenshot guardado: error_ficha_{ficha}_form.png")
            except:
                pass
            return False

    def click_consultar_button_in_iframe(self):
        """Hace clic en el botón 'Consultar' dentro del iframe para ejecutar la búsqueda"""
        try:
            print("Buscando el botón 'Consultar' (form:buscarCBT) dentro del iframe...")
            
            # Estrategia 1: Buscar por ID específico
            try:
                consultar_button = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "form:buscarCBT"))
                )
                print("✓ Botón 'Consultar' encontrado por ID en iframe")
                
                # Hacer scroll al botón
                self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                time.sleep(0.5)
                
                # Hacer clic
                consultar_button.click()
                print("✓ Clic exitoso en el botón 'Consultar' en iframe")
                
                # Esperar a que se procese la búsqueda
                time.sleep(3)
                return True
                
            except:
                print("❌ No se encontró el botón por ID en iframe, probando otros selectores...")
                
                # Estrategia 2: Buscar por name
                try:
                    consultar_button = self.driver.find_element(By.NAME, "form:buscarCBT")
                    if consultar_button.is_displayed() and consultar_button.is_enabled():
                        print("✓ Botón 'Consultar' encontrado por NAME en iframe")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                        time.sleep(0.5)
                        consultar_button.click()
                        print("✓ Clic exitoso en el botón 'Consultar' en iframe")
                        time.sleep(3)
                        return True
                except:
                    pass
                
                # Estrategia 3: Buscar por valor "Consultar"
                try:
                    consultar_button = self.driver.find_element(By.XPATH, "//input[@value='Consultar']")
                    if consultar_button.is_displayed() and consultar_button.is_enabled():
                        print("✓ Botón 'Consultar' encontrado por valor en iframe")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                        time.sleep(0.5)
                        consultar_button.click()
                        print("✓ Clic exitoso en el botón 'Consultar' en iframe")
                        time.sleep(3)
                        return True
                except:
                    pass
            
            print("❌ No se pudo encontrar el botón 'Consultar' con ninguna estrategia en iframe")
            return False
            
        except Exception as e:
            print(f"❌ Error al buscar el botón 'Consultar' en iframe: {e}")
            return False

    def click_agregar_button(self):
        """Hace clic en el botón 'Agregar' en los resultados de la búsqueda (dentro del iframe)"""
        try:
            print("Buscando el botón 'Agregar' en los resultados (en iframe)...")
            
            # Estrategia 1: Buscar por ID específico (primer resultado)
            try:
                agregar_button = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "form:dtFichas:0:imgSelec"))
                )
                print("✓ Botón 'Agregar' encontrado por ID específico (form:dtFichas:0:imgSelec) en iframe")
                
                # Hacer scroll al botón
                self.driver.execute_script("arguments[0].scrollIntoView(true);", agregar_button)
                time.sleep(0.5)
                
                # Hacer clic
                agregar_button.click()
                print("✓ Clic exitoso en el botón 'Agregar' en iframe")
                
                # Esperar a que se procese la acción
                time.sleep(3)
                return True
                
            except:
                print("❌ No se encontró el botón por ID específico en iframe, probando patrones generales...")
                
                # Estrategia 2: Buscar por patrón de ID (cualquier fila)
                try:
                    agregar_buttons = self.driver.find_elements(By.XPATH, "//img[contains(@id, 'dtFichas') and contains(@id, 'imgSelec')]")
                    if agregar_buttons:
                        agregar_button = agregar_buttons[0]  # Tomar el primero
                        button_id = agregar_button.get_attribute('id')
                        print(f"✓ Botón 'Agregar' encontrado por patrón: {button_id} en iframe")
                        
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", agregar_button)
                        time.sleep(0.5)
                        agregar_button.click()
                        print("✓ Clic exitoso en el botón 'Agregar' en iframe")
                        time.sleep(3)
                        return True
                except:
                    pass
                
                # Estrategia 3: Buscar por title "Agregar"
                try:
                    agregar_button = self.driver.find_element(By.XPATH, "//img[@title='Agregar']")
                    if agregar_button.is_displayed():
                        print("✓ Botón 'Agregar' encontrado por title en iframe")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", agregar_button)
                        time.sleep(0.5)
                        agregar_button.click()
                        print("✓ Clic exitoso en el botón 'Agregar' en iframe")
                        time.sleep(3)
                        return True
                except:
                    pass
            
            print("❌ No se pudo encontrar el botón 'Agregar' con ninguna estrategia en iframe")
            return False
            
        except Exception as e:
            print(f"❌ Error al buscar el botón 'Agregar' en iframe: {e}")
            return False


    def find_and_process_input_in_all_contexts(self, ficha):
        """Busca el campo de input en todos los contextos posibles (fallback)"""
        try:
            print("Buscando el campo de input en todos los contextos (fallback)...")
            
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
                time.sleep(1)
                return True
            else:
                print(f"❌ Error: Se esperaba '{ficha}' pero se insertó '{inserted_value}'")
                return False
                
        except Exception as e:
            print(f"❌ Error al procesar el campo de input: {e}")
            return False

    def process_single_ficha(self, ficha):
        """Procesa una sola ficha: hace clic en el botón, inserta la ficha, ejecuta la consulta, agrega el resultado y selecciona la opción en el iframe."""
        try:
            print(f"\n{'='*50}")
            print(f"PROCESANDO FICHA: {ficha}")
            print(f"{'='*50}")
            
            # Paso 1: Hacer clic en el botón "Consultar ficha" (esto abre el iframe)
            if not self.click_consultar_ficha_button():
                print(f"❌ No se pudo hacer clic en 'Consultar ficha' para la ficha {ficha}")
                return False
            
            # Paso 2: Esperar el formulario, insertar la ficha, hacer clic en "Consultar", "Agregar" y seleccionar opción
            # TODO ESTO OCURRE DENTRO DEL IFRAME
            if not self.wait_for_form_and_insert_ficha(ficha):
                print(f"❌ No se pudo completar el proceso para la ficha {ficha}")
                return False
            
            print(f"✓ Ficha {ficha} procesada exitosamente y 'Primera Opción' seleccionada en el iframe")
            
            # Después de procesar la ficha y seleccionar la opción, la modal/iframe debería cerrarse automáticamente
            # o el siguiente paso nos llevará de vuelta al contexto principal.
            # Por ahora, no hacemos switch_to.default_content() aquí, se hará antes de click_consultar_aspirantes_button
            
            return True
            
        except Exception as e:
            print(f"❌ Error al procesar ficha {ficha}: {e}")
            return False

    def click_consultar_aspirantes_button(self):
        """Hace clic en el botón 'Consultar aspirantes' después de procesar todas las fichas"""
        try:
            print("Regresando al contexto principal para buscar 'Consultar aspirantes'...")
            self.driver.switch_to.default_content()
            time.sleep(2)
            
            print("Buscando el botón 'Consultar aspirantes' en la página principal...")
            
            # Estrategia 1: Buscar por ID específico
            try:
                consultar_button = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "frmPrincipal:cmdlnkSearch"))
                )
                print("✓ Botón 'Consultar aspirantes' encontrado por ID")
                
                # Hacer scroll al botón
                self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                time.sleep(0.5)
                
                # Hacer clic
                consultar_button.click()
                print("✓ Clic exitoso en el botón 'Consultar aspirantes'")
                
                # Esperar a que se genere el reporte
                time.sleep(5)
                return True
                
            except:
                print("❌ No se encontró el botón por ID, probando otros selectores...")
                
                # Estrategia 2: Buscar por valor "Consultar aspirantes"
                try:
                    consultar_button = self.driver.find_element(By.XPATH, "//input[@value='Consultar aspirantes ']")
                    if consultar_button.is_displayed() and consultar_button.is_enabled():
                        print("✓ Botón 'Consultar aspirantes' encontrado por valor")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                        time.sleep(0.5)
                        consultar_button.click()
                        print("✓ Clic exitoso en el botón 'Consultar aspirantes'")
                        time.sleep(5)
                        return True
                except:
                    pass
                
                # Estrategia 3: Buscar por cualquier botón que contenga "Consultar"
                try:
                    consultar_button = self.driver.find_element(By.XPATH, "//input[contains(@value, 'Consultar')]")
                    if consultar_button.is_displayed() and consultar_button.is_enabled():
                        print("✓ Botón 'Consultar' encontrado por contenido")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                        time.sleep(0.5)
                        consultar_button.click()
                        print("✓ Clic exitoso en el botón 'Consultar'")
                        time.sleep(5)
                        return True
                except:
                    pass
            
            print("❌ No se pudo encontrar el botón 'Consultar aspirantes' con ninguna estrategia")
            return False
            
        except Exception as e:
            print(f"❌ Error al buscar el botón 'Consultar aspirantes': {e}")
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
            
            # NO SELECCIONAR LA OPCIÓN INICIALMENTE EN LA PÁGINA PRINCIPAL
            # if not self.select_primera_opcion_initial():
            #     print("❌ Error al configurar 'Primera Opción' inicialmente")
            #     return
            
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
                    time.sleep(5)  # Tiempo para que se procesen los resultados
            
            # Después de procesar todas las fichas, hacer clic en "Consultar aspirantes"
            if successful > 0:
                print(f"\n{'='*50}")
                print(f"GENERANDO REPORTE FINAL")
                print(f"{'='*50}")
                
                if self.click_consultar_aspirantes_button():
                    print("✓ Reporte generado exitosamente")
                else:
                    print("❌ Error al generar el reporte final")
            
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
