import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import time
import os
from selenium.webdriver.common.action_chains import ActionChains
import logging
from datetime import datetime
from pathlib import Path
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.firefox.options import Options

class SenaAutomation:
    def __init__(self):
        self.setup_logging()
        self.setup_driver()
        
    def setup_logging(self):
        """Configura el sistema de logging avanzado"""
        # Crear directorio de logs
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # Configurar logging con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"sena_automation_{timestamp}.log")
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.html_dump_file = os.path.join(log_dir, f"html_dumps_{timestamp}.html")
        
        # Inicializar archivo HTML
        with open(self.html_dump_file, 'w', encoding='utf-8') as f:
            f.write("<html><head><title>SENA Automation HTML Dumps</title></head><body>\n")
        
        self.logger.info("Sistema de logging inicializado")
            

    def setup_driver(self):
        """Configura el driver de Firefox para descargar sin preguntar y guardar en Descargas."""

        # Carpeta Descargas
        download_dir = str(Path.home() / "Downloads")

        # Crear perfil
        profile = FirefoxProfile()

        # Configuración de descarga automática
        profile.set_preference("browser.download.folderList", 2)
        profile.set_preference("browser.download.dir", download_dir)
        profile.set_preference("browser.download.useDownloadDir", True)
        profile.set_preference("browser.download.manager.showWhenStarting", False)
        profile.set_preference("browser.download.manager.alertOnEXEOpen", False)

        # Tipos de archivo para descarga directa
        mime_types = [
            "application/pdf",
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/msword",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "text/csv"
        ]
        profile.set_preference("browser.helperApps.neverAsk.saveToDisk", ",".join(mime_types))
        profile.set_preference("pdfjs.disabled", True)

        # Configurar opciones y asignar perfil
        options = Options()
        options.profile = profile

        # Iniciar Firefox con el perfil
        self.driver = webdriver.Firefox(options=options)
        self.wait = WebDriverWait(self.driver, 20)
        self.logger.info(f"Firefox configurado para descargar en: {download_dir}")
        

    def save_page_html(self, step_name, additional_info=""):
        """Guarda el HTML completo de la página actual con análisis"""
        try:
            timestamp = datetime.now().strftime("%H:%M:%S")
            
            # Obtener HTML de la página principal
            main_html = self.driver.page_source
            
            # Analizar elementos importantes
            analysis = self.analyze_page_elements()
            
            # Guardar en archivo HTML
            with open(self.html_dump_file, 'a', encoding='utf-8') as f:
                f.write(f"\n<hr><h2>PASO: {step_name} - {timestamp}</h2>\n")
                if additional_info:
                    f.write(f"<p><strong>Info adicional:</strong> {additional_info}</p>\n")
                
                f.write(f"<h3>Análisis de elementos:</h3>\n<pre>{analysis}</pre>\n")
                f.write(f"<h3>HTML completo:</h3>\n<textarea style='width:100%;height:400px;'>{main_html}</textarea>\n")
            
            self.logger.info(f"HTML guardado para paso: {step_name}")
            
        except Exception as e:
            self.logger.error(f"Error guardando HTML: {e}")

    def analyze_page_elements(self):
        """Analiza todos los elementos importantes de la página"""
        try:
            analysis = []
            analysis.append("="*80)
            analysis.append("ANÁLISIS COMPLETO DE ELEMENTOS")
            analysis.append("="*80)
            
            # Información básica
            analysis.append(f"URL actual: {self.driver.current_url}")
            analysis.append(f"Título: {self.driver.title}")
            
            # Buscar todos los botones
            analysis.append("\n--- TODOS LOS BOTONES ---")
            buttons = self.driver.find_elements(By.TAG_NAME, "button") + \
                     self.driver.find_elements(By.CSS_SELECTOR, "input[type='button']") + \
                     self.driver.find_elements(By.CSS_SELECTOR, "input[type='submit']")
            
            for i, btn in enumerate(buttons):
                try:
                    btn_id = btn.get_attribute('id') or 'Sin ID'
                    btn_name = btn.get_attribute('name') or 'Sin name'
                    btn_value = btn.get_attribute('value') or 'Sin value'
                    btn_text = btn.text or 'Sin texto'
                    btn_class = btn.get_attribute('class') or 'Sin class'
                    btn_onclick = btn.get_attribute('onclick') or 'Sin onclick'
                    btn_visible = btn.is_displayed()
                    btn_enabled = btn.is_enabled()
                    
                    analysis.append(f"Botón {i+1}:")
                    analysis.append(f"  ID: {btn_id}")
                    analysis.append(f"  Name: {btn_name}")
                    analysis.append(f"  Value: {btn_value}")
                    analysis.append(f"  Text: {btn_text}")
                    analysis.append(f"  Class: {btn_class}")
                    analysis.append(f"  OnClick: {btn_onclick}")
                    analysis.append(f"  Visible: {btn_visible}, Enabled: {btn_enabled}")
                    analysis.append("")
                except:
                    analysis.append(f"Botón {i+1}: Error al obtener información")
            
            # Buscar elementos con "consultar" en el texto
            analysis.append("\n--- ELEMENTOS CON 'CONSULTAR' ---")
            consultar_elements = self.driver.find_elements(By.XPATH, "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'consultar')]")
            
            for i, elem in enumerate(consultar_elements):
                try:
                    elem_tag = elem.tag_name
                    elem_id = elem.get_attribute('id') or 'Sin ID'
                    elem_text = elem.text or 'Sin texto'
                    elem_visible = elem.is_displayed()
                    
                    analysis.append(f"Elemento {i+1}: <{elem_tag}> ID: {elem_id}, Text: '{elem_text}', Visible: {elem_visible}")
                except:
                    analysis.append(f"Elemento {i+1}: Error al obtener información")
            
            # Buscar formularios
            analysis.append("\n--- FORMULARIOS ---")
            forms = self.driver.find_elements(By.TAG_NAME, "form")
            for i, form in enumerate(forms):
                try:
                    form_id = form.get_attribute('id') or 'Sin ID'
                    form_name = form.get_attribute('name') or 'Sin name'
                    form_action = form.get_attribute('action') or 'Sin action'
                    
                    analysis.append(f"Formulario {i+1}: ID: {form_id}, Name: {form_name}, Action: {form_action}")
                except:
                    analysis.append(f"Formulario {i+1}: Error al obtener información")
            
            return "\n".join(analysis)
            
        except Exception as e:
            return f"Error en análisis: {e}"

    def validate_iframes_after_modal_close(self):
        """Valida el iframe específico 'contenido' después de cerrar la modal"""
        try:
            target_iframe = None
            
            # Buscar por ID primero
            try:
                target_iframe = self.driver.find_element(By.ID, "contenido")
            except:
                # Si no se encuentra por ID, buscar por name
                try:
                    target_iframe = self.driver.find_element(By.NAME, "contenido")
                except:
                    # Buscar por src que contenga la ruta específica
                    try:
                        target_iframe = self.driver.find_element(
                            By.XPATH, 
                            "//iframe[contains(@src, '/sofia/inscripcion/generarreporteinscripcion/generarReporteInscripcion.faces')]"
                        )
                    except:
                        return [], []
            
            if not target_iframe:
                return [], []
            
            # Obtener información del iframe encontrado
            iframe_id = target_iframe.get_attribute('id') or 'Sin ID'
            iframe_name = target_iframe.get_attribute('name') or 'Sin name'
            iframe_src = target_iframe.get_attribute('src') or 'Sin src'
            
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame(target_iframe)
            
            # Buscar el botón directamente en este iframe
            button_found = self.search_button_in_specific_iframe()
            
            iframe_info = [{
                'index': 0,
                'id': iframe_id,
                'name': iframe_name,
                'src': iframe_src,
                'target': True
            }]
            
            if button_found:
                return iframe_info, [0]
            else:
                return iframe_info, []
                
        except Exception as e:
            return [], []

    def search_button_in_specific_iframe(self):
        """Busca el botón 'Consultar aspirantes' en el iframe actual"""
        try:
            # Selectores para el botón
            selectors = [
                (By.ID, "frmPrincipal:cmdlnkSearch"),
                (By.NAME, "frmPrincipal:cmdlnkSearch"),
                (By.XPATH, "//input[@value='Consultar aspirantes ']"),
                (By.XPATH, "//input[contains(@value, 'Consultar aspirantes')]"),
                (By.CLASS_NAME, "boton_app")
            ]
            
            for selector_type, selector_value in selectors:
                try:
                    element = self.driver.find_element(selector_type, selector_value)
                    if element and element.is_displayed():
                        return True
                except:
                    continue
            
            return False
            
        except Exception as e:
            return False

    def click_consultar_aspirantes_with_validation(self):
        """Hace clic en 'Consultar aspirantes' en el iframe 'contenido'"""
        try:
            self.driver.switch_to.default_content()
            
            # Buscar el iframe "contenido"
            target_iframe = None
            try:
                target_iframe = self.driver.find_element(By.ID, "contenido")
            except:
                try:
                    target_iframe = self.driver.find_element(By.NAME, "contenido")
                except:
                    try:
                        target_iframe = self.driver.find_element(
                            By.XPATH, 
                            "//iframe[contains(@src, '/sofia/inscripcion/generarreporteinscripcion/generarReporteInscripcion.faces')]"
                        )
                    except:
                        return False
            
            if not target_iframe:
                return False
            
            # Cambiar al iframe específico
            self.driver.switch_to.frame(target_iframe)
            
            # Esperar que el botón esté disponible
            wait = WebDriverWait(self.driver, 15)
            
            # Selectores para el botón
            selectors = [
                (By.ID, "frmPrincipal:cmdlnkSearch"),
                (By.NAME, "frmPrincipal:cmdlnkSearch"),
                (By.XPATH, "//input[@value='Consultar aspirantes ']"),
                (By.XPATH, "//input[contains(@value, 'Consultar aspirantes')]"),
                (By.CLASS_NAME, "boton_app")
            ]
            
            boton_aspirantes = None
            for selector_type, selector_value in selectors:
                try:
                    boton_aspirantes = wait.until(
                        EC.element_to_be_clickable((selector_type, selector_value))
                    )
                    break
                except:
                    continue
            
            if not boton_aspirantes:
                return False
            
            # Scroll al elemento
            self.driver.execute_script("arguments[0].scrollIntoView(true);", boton_aspirantes)
            time.sleep(1)
            
            click_strategies = [
                lambda: boton_aspirantes.click(),
                lambda: self.driver.execute_script("arguments[0].click();", boton_aspirantes),
                lambda: self.driver.execute_script("document.forms['frmPrincipal'].submit();"),
                lambda: ActionChains(self.driver).move_to_element(boton_aspirantes).click().perform()
            ]
            
            for i, strategy in enumerate(click_strategies, 1):
                try:
                    strategy()
                    time.sleep(2)
                    
                    # Verificar éxito básico
                    current_url = self.driver.current_url
                    if "aspirantes" in current_url.lower() or "consulta" in current_url.lower():
                        return True
                        
                except Exception as e:
                    continue
            
            return False
            
        except Exception as e:
            return False


    def execute_multiple_click_strategies(self):
        """Ejecuta múltiples estrategias de clic en el botón"""
        try:
            # Encontrar el botón
            button = None
            selectors = [
                (By.ID, "frmPrincipal:cmdlnkSearch"),
                (By.NAME, "frmPrincipal:cmdlnkSearch"),
                (By.XPATH, "//input[@value='Consultar aspirantes ']")
            ]
            
            for by_type, selector in selectors:
                try:
                    button = self.driver.find_element(by_type, selector)
                    if button.is_displayed():
                        break
                except:
                    continue
            
            if not button:
                self.logger.error("❌ No se pudo encontrar el botón para hacer clic")
                return False
            
            # Guardar HTML antes del clic
            self.save_page_html("Antes_del_clic_consultar_aspirantes", f"Botón encontrado: {button.get_attribute('id')}")
            
            # Estrategias de clic
            strategies = [
                ("Click normal", lambda: button.click()),
                ("JavaScript click", lambda: self.driver.execute_script("arguments[0].click();", button)),
                ("ActionChains click", lambda: ActionChains(self.driver).click(button).perform()),
                ("Submit form", lambda: self.driver.execute_script("arguments[0].form.submit();", button)),
                ("Trigger onclick", lambda: self.driver.execute_script("arguments[0].onclick();", button) if button.get_attribute('onclick') else None),
                ("Focus + Enter", lambda: (button.click(), self.driver.execute_script("arguments[0].focus();", button))),
                ("Dispatch Event", lambda: self.driver.execute_script("""
                    var event = new MouseEvent('click', {
                        view: window,
                        bubbles: true,
                        cancelable: true
                    });
                    arguments[0].dispatchEvent(event);
                """, button))
            ]
            
            for strategy_name, strategy_func in strategies:
                try:
                    self.logger.info(f"🔄 Ejecutando estrategia: {strategy_name}")
                    
                    # Hacer scroll al botón
                    self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", button)
                    time.sleep(1)
                    
                    # Resaltar el botón
                    original_style = button.get_attribute('style')
                    self.driver.execute_script("arguments[0].style.border='3px solid red';", button)
                    time.sleep(0.5)
                    
                    # Ejecutar estrategia
                    strategy_func()
                    
                    # Restaurar estilo
                    self.driver.execute_script(f"arguments[0].style='{original_style}';", button)
                    
                    self.logger.info(f"✅ Estrategia '{strategy_name}' ejecutada")
                    
                    # Verificar éxito
                    if self.verify_click_success():
                        self.logger.info(f"🎉 Clic exitoso con estrategia: {strategy_name}")
                        return True
                    
                    time.sleep(2)  # Pausa entre estrategias
                    
                except Exception as e:
                    self.logger.error(f"❌ Error con estrategia '{strategy_name}': {e}")
                    continue
            
            self.logger.warning("⚠️ Todas las estrategias ejecutadas, verificar manualmente")
            return True  # Asumir éxito parcial
            
        except Exception as e:
            self.logger.error(f"❌ Error ejecutando estrategias de clic: {e}")
            return False

    def verify_click_success(self):
        """Verifica si el clic fue exitoso"""
        try:
            # Esperar cambios en la página
            time.sleep(3)
            
            # Verificar si hay indicadores de carga o cambios
            loading_indicators = [
                "//div[contains(@class, 'loading')]",
                "//div[contains(@class, 'spinner')]",
                "//*[contains(text(), 'Cargando')]",
                "//*[contains(text(), 'Procesando')]"
            ]
            
            for indicator in loading_indicators:
                try:
                    if self.driver.find_elements(By.XPATH, indicator):
                        self.logger.info("✅ Indicador de carga detectado - clic exitoso")
                        return True
                except:
                    continue
            
            # Verificar cambios en el DOM
            try:
                WebDriverWait(self.driver, 5).until(
                    lambda driver: driver.execute_script("return document.readyState") == "complete"
                )
                return True
            except:
                pass
            
            return False
            
        except Exception as e:
            self.logger.error(f"Error verificando éxito del clic: {e}")
            return False

    def debug_page_state_after_modal(self):
        """Debug completo del estado de la página después de cerrar modal"""
        try:
            self.logger.info("🔍 DEBUG COMPLETO DEL ESTADO DE LA PÁGINA")
            
            # Guardar HTML completo
            self.save_page_html("Debug_estado_completo_post_modal")
            
            # Validar iframes
            self.validate_iframes_after_modal_close()
            
            # Verificar errores JavaScript
            try:
                js_errors = self.driver.execute_script("return window.jsErrors || [];")
                if js_errors:
                    self.logger.warning(f"⚠️ Errores JavaScript detectados: {js_errors}")
                else:
                    self.logger.info("✅ No hay errores JavaScript")
            except:
                self.logger.info("No se pudieron verificar errores JavaScript")
            
            # Estado del documento
            try:
                ready_state = self.driver.execute_script("return document.readyState")
                self.logger.info(f"Estado del documento: {ready_state}")
            except:
                pass
            
            self.logger.info("🔍 Debug completo finalizado - revisar logs y archivos HTML")
            
        except Exception as e:
            self.logger.error(f"Error en debug: {e}")

            

    def click_agregar_y_consultar_aspirantes_with_validation(self, timeout=10):
        """Clic en Agregar, luego en Consultar Aspirantes y finalmente en Generar Reporte."""
        try:
            self.logger.info("🎯 AGREGAR Y CONSULTAR ASPIRANTES CON VALIDACIÓN")

            # Paso 1: Clic en Agregar
            agregar_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "form:dtFichas:0:imgSelec"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", agregar_button)
            time.sleep(0.5)
            agregar_button.click()
            self.logger.info("✅ Clic en 'Agregar' exitoso")

            # Paso 2: Esperar iframe objetivo
            self.driver.switch_to.default_content()
            self.logger.info("⏳ Esperando iframe 'contenido'...")
            target_iframe = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//iframe[@id='contenido' or @name='contenido' or contains(@src, 'generarReporteInscripcion.faces')]")
                )
            )
            self.driver.switch_to.frame(target_iframe)
            self.logger.info("✅ Iframe 'contenido' encontrado y activado")

            # Paso 3: Clic en Consultar Aspirantes
            consultar_btn = WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, "//input[contains(@value, 'Consultar aspirantes')]"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_btn)
            time.sleep(0.5)
            consultar_btn.click()
            self.logger.info("✅ Clic en 'Consultar aspirantes' exitoso")

            # Paso 4: Esperar y hacer clic en Generar Reporte
            self.logger.info("⏳ Esperando 2 segundos para 'Generar reporte'...")
            time.sleep(2)
            generar_btn = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.ID, "frmPrincipal:btnGenerar"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", generar_btn)
            time.sleep(0.5)
            generar_btn.click()
            self.logger.info("✅ Clic en 'Generar reporte' exitoso")

            # Volver al contexto principal
            self.driver.switch_to.default_content()
            return True

        except Exception as e:
            self.logger.error(f"❌ Error en flujo completo: {e}")
            self.driver.switch_to.default_content()
            return False



    # ... resto de métodos originales sin cambios ...
    def read_excel_fichas(self, excel_path):
        """Lee el archivo Excel y extrae los números de ficha"""
        try:
            df = pd.read_excel(excel_path)
            fichas = df.iloc[:, 0].tolist()
            self.logger.info(f"Se encontraron {len(fichas)} fichas en el Excel")
            return fichas
        except Exception as e:
            self.logger.error(f"Error al leer el Excel: {e}")
            return []

    def switch_to_login_iframe(self):
        """Cambia al iframe donde está el formulario de login"""
        try:
            self.logger.info("Buscando iframe de login...")
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
            
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            self.logger.info(f"Se encontraron {len(iframes)} iframes")
            
            for i, iframe in enumerate(iframes):
                try:
                    self.logger.info(f"Probando iframe {i+1}...")
                    self.driver.switch_to.frame(iframe)
                    
                    try:
                        self.driver.find_element(By.ID, "username")
                        self.logger.info(f"¡Formulario de login encontrado en iframe {i+1}!")
                        return True
                    except:
                        self.driver.switch_to.default_content()
                        continue
                        
                except Exception as e:
                    self.logger.error(f"Error al acceder al iframe {i+1}: {e}")
                    self.driver.switch_to.default_content()
                    continue
            
            self.logger.error("No se encontró el iframe con el formulario de login")
            return False
            
        except Exception as e:
            self.logger.error(f"Error al buscar iframes: {e}")
            return False

    def navigate_to_sena(self):
        """Navega al sitio de SENA Sofia Plus y hace login"""
        try:
            self.logger.info("Navegando a SENA Sofia Plus...")
            self.driver.get("http://senasofiaplus.edu.co/sofia-public/")
            time.sleep(3)
            
            try:
                continue_button = self.driver.find_element(By.ID, "proceed-button")
                continue_button.click()
                self.logger.info("Advertencia SSL manejada")
                time.sleep(2)
            except:
                self.logger.info("No se encontró advertencia SSL o ya se pasó")
            
            if not self.switch_to_login_iframe():
                return False
            
            if not self.login():
                return False
            
            if not self.post_login_navigation():
                return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error al navegar al sitio: {e}")
            return False

    def login(self):
        """Realiza el proceso de login"""
        try:
            self.logger.info("Iniciando proceso de login...")
            self.wait.until(EC.presence_of_element_located((By.ID, "username")))
            
            documento_field = self.driver.find_element(By.ID, "username")
            documento_field.clear()
            documento_field.send_keys("38363049")
            self.logger.info("Número de documento ingresado")
            
            password_field = self.driver.find_element(By.NAME, "josso_password")
            password_field.clear()
            password_field.send_keys("10337684990Ad*")
            self.logger.info("Contraseña ingresada")
            
            ingresar_button = self.driver.find_element(By.CSS_SELECTOR, "input.login100-form-btn[value='Ingresar']")
            ingresar_button.click()
            self.logger.info("Botón INGRESAR presionado")
            
            time.sleep(5)
            self.driver.switch_to.default_content()
            self.logger.info("Volviendo al contexto principal")
            
            try:
                time.sleep(3)
                current_url = self.driver.current_url
                self.logger.info(f"URL actual después del login: {current_url}")
                
                if "login" not in current_url.lower() or len(self.driver.find_elements(By.TAG_NAME, "iframe")) > 0:
                    self.logger.info("Login exitoso")
                    return True
                else:
                    self.logger.error("Login falló - aún en página de login")
                    return False
                    
            except Exception as e:
                self.logger.error(f"Error verificando login: {e}")
                return False
                
        except Exception as e:
            self.logger.error(f"Error durante el login: {e}")
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
            self.logger.error(f"Error en la navegación post-login: {e}")
            return False

    def select_role(self):
        """Selecciona el rol después del login"""
        try:
            self.logger.info("Seleccionando rol...")
            role_select = self.wait.until(
                EC.presence_of_element_located((By.ID, "seleccionRol:roles"))
            )
            
            select = Select(role_select)
            select.select_by_value("33")
            self.logger.info("Rol 'Encargado de ingreso centro formación' seleccionado")
            time.sleep(3)
            return True
            
        except Exception as e:
            self.logger.error(f"Error al seleccionar rol: {e}")
            return False

    def navigate_to_inscripcion(self):
        """Navega a la opción 'Inscripción' en el menú principal"""
        try:
            self.logger.info("Navegando a 'Inscripción'...")
            
            inscripcion_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[@class='menuPrimario' and text()='Inscripción']"))
            )
            inscripcion_link.click()
            self.logger.info("Clic en 'Inscripción' exitoso")
            time.sleep(3)
            
            consultas_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Consultas')]"))
            )
            consultas_link.click()
            self.logger.info("Clic en 'Consultas' exitoso")
            time.sleep(2)
            
            generar_reporte_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Generar Reporte de Inscripción')]"))
            )
            generar_reporte_link.click()
            self.logger.info("Clic en 'Generar Reporte de Inscripción' exitoso")
            time.sleep(3)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error al navegar a 'Inscripción': {e}")
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
                    self.logger.info("✓ 'Primera Opción' seleccionada correctamente")
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
                            self.logger.info("✓ 'Primera Opción' seleccionada correctamente (iframe interno)")
                            return [i, j]  # Ruta: externo + interno
                        except:
                            continue

            self.logger.error("❌ No se encontró el select 'opcionesInscritos'")
            return None

        except Exception as e:
            self.logger.error(f"❌ Error seleccionando opción: {e}")
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
                    self.logger.info(f"✓ Clic exitoso con selector: {selector}")
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
                self.logger.error("❌ No se pudo seleccionar la opción.")
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
                self.logger.info("✓ Clic exitoso en 'Consultar ficha'")
                return True
            else:
                self.logger.error("❌ No se encontró el botón 'Consultar ficha' en el mismo iframe")
                return False

        except Exception as e:
            self.logger.error(f"❌ Error al buscar el botón: {e}")
            return False

    def wait_for_form_and_insert_ficha(self, ficha):
        """Espera a que aparezca el formulario y procesa la ficha completa"""
        try:
            self.logger.info(f"Esperando formulario para ficha: {ficha}")
            
            # PASO 1: Esperar un poco para que el formulario se cargue después del clic
            self.logger.info("Esperando que se cargue el formulario...")
            time.sleep(3)
            
            # PASO 2: Buscar e insertar la ficha en el campo de input
            self.logger.info("Buscando el campo de input en el contexto actual (iframe)...")
            input_ficha = None
            
            try:
                # Intentar encontrar el input en el iframe actual (donde se hizo el clic)
                input_ficha = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "form:codigoFichaITX"))
                )
                self.logger.info("✓ Campo de input encontrado en el iframe actual")
                
                # Procesar el input
                if not self.process_input_field(input_ficha, ficha):
                    return False
                
            except:
                self.logger.error("❌ No se encontró el input en el iframe actual, buscando en otros contextos...")
                # Si no se encuentra en el iframe actual, buscar en todos los contextos
                if not self.find_and_process_input_in_all_contexts(ficha):
                    return False
            
            # PASO 3: Hacer clic en el botón "Consultar" después de insertar la ficha
            self.logger.info("Buscando el botón 'Consultar' para ejecutar la búsqueda (en iframe)...")
            if not self.click_consultar_button_in_iframe():
                self.logger.error("❌ No se pudo hacer clic en el botón 'Consultar'")
                return False
            
            # PASO 4: Esperar 2 segundos y hacer clic en el botón "Agregar"
            self.logger.info("Esperando 2 segundos para que aparezcan los resultados (en iframe)...")
            time.sleep(2)


            
            # USAR LA NUEVA FUNCIÓN CON VALIDACIÓN
            if not self.click_agregar_y_consultar_aspirantes_with_validation():
                self.logger.error(f"❌ No se pudo completar el paso Agregar -> Consultar aspirantes para ficha {ficha}")
                return False

            return True
                
        except Exception as e:
            self.logger.error(f"❌ Error general al procesar ficha {ficha}: {e}")
            try:
                self.driver.save_screenshot(f"error_ficha_{ficha}_form.png")
                self.logger.info(f"Screenshot guardado: error_ficha_{ficha}_form.png")
            except:
                pass
            return False

    def click_consultar_button_in_iframe(self):
        """Hace clic en el botón 'Consultar' dentro del iframe para ejecutar la búsqueda"""
        try:
            self.logger.info("Buscando el botón 'Consultar' (form:buscarCBT) dentro del iframe...")
            
            # Estrategia 1: Buscar por ID específico
            try:
                consultar_button = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "form:buscarCBT"))
                )
                self.logger.info("✓ Botón 'Consultar' encontrado por ID en iframe")
                
                # Hacer scroll al botón
                self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                time.sleep(0.5)
                
                # Hacer clic
                consultar_button.click()
                self.logger.info("✓ Clic exitoso en el botón 'Consultar' en iframe")
                
                # Esperar a que se procese la búsqueda
                self.logger.info("Esperando 2 segundos para que aparezcan los resultados (en iframe)...")
                time.sleep(2)
                return True
                
            except:
                self.logger.error("❌ No se encontró el botón por ID en iframe, probando otros selectores...")
                
                # Estrategia 2: Buscar por name
                try:
                    consultar_button = self.driver.find_element(By.NAME, "form:buscarCBT")
                    if consultar_button.is_displayed() and consultar_button.is_enabled():
                        self.logger.info("✓ Botón 'Consultar' encontrado por NAME en iframe")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                        time.sleep(0.5)
                        consultar_button.click()
                        self.logger.info("✓ Clic exitoso en el botón 'Consultar' en iframe")
                        time.sleep(2)
                        return True
                except:
                    pass
                
                # Estrategia 3: Buscar por valor "Consultar"
                try:
                    consultar_button = self.driver.find_element(By.XPATH, "//input[@value='Consultar']")
                    if consultar_button.is_displayed() and consultar_button.is_enabled():
                        self.logger.info("✓ Botón 'Consultar' encontrado por valor en iframe")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", consultar_button)
                        time.sleep(0.5)
                        consultar_button.click()
                        self.logger.info("✓ Clic exitoso en el botón 'Consultar' en iframe")
                        time.sleep(2)
                        return True
                except:
                    pass
            
            self.logger.error("❌ No se pudo encontrar el botón 'Consultar' con ninguna estrategia en iframe")
            return False
            
        except Exception as e:
            self.logger.error(f"❌ Error al buscar el botón 'Consultar' en iframe: {e}")
            return False

    def find_and_process_input_in_all_contexts(self, ficha):
        """Busca el campo de input en todos los contextos posibles (fallback)"""
        try:
            self.logger.info("Buscando el campo de input en todos los contextos (fallback)...")
            
            # Contexto 1: Página principal
            self.logger.info("Probando contexto principal...")
            self.driver.switch_to.default_content()
            try:
                input_ficha = self.driver.find_element(By.ID, "form:codigoFichaITX")
                if input_ficha.is_displayed():
                    self.logger.info("✓ Campo encontrado en contexto principal")
                    return self.process_input_field(input_ficha, ficha)
            except:
                pass
            
            # Contexto 2: Buscar en todos los iframes
            self.logger.info("Buscando en todos los iframes...")
            self.driver.switch_to.default_content()
            all_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            
            for i, iframe in enumerate(all_iframes):
                try:
                    iframe_id = iframe.get_attribute('id')
                    self.logger.info(f"Probando iframe {i+1} (ID: {iframe_id})...")
                    
                    # Cambiar al iframe
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame(iframe)
                    
                    # Buscar el input
                    try:
                        input_ficha = self.driver.find_element(By.ID, "form:codigoFichaITX")
                        if input_ficha.is_displayed():
                            self.logger.info(f"✓ Campo encontrado en iframe {i+1}")
                            return self.process_input_field(input_ficha, ficha)
                    except:
                        pass
                    
                    # Buscar en iframes anidados
                    nested_iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                    for j, nested in enumerate(nested_iframes):
                        try:
                            self.logger.info(f"  Probando iframe anidado {i+1}.{j+1}...")
                            self.driver.switch_to.default_content()
                            self.driver.switch_to.frame(iframe)
                            self.driver.switch_to.frame(nested)
                            
                            input_ficha = self.driver.find_element(By.ID, "form:codigoFichaITX")
                            if input_ficha.is_displayed():
                                self.logger.info(f"✓ Campo encontrado en iframe anidado {i+1}.{j+1}")
                                return self.process_input_field(input_ficha, ficha)
                                
                        except:
                            continue
                            
                except Exception as e:
                    continue
            
            self.logger.error("❌ No se encontró el campo de input en ningún contexto")
            self.driver.switch_to.default_content()
            return False
            
        except Exception as e:
            self.logger.error(f"Error buscando en contextos: {e}")
            self.driver.switch_to.default_content()
            return False

    def process_input_field(self, input_element, ficha):
        """Procesa el campo de input insertando la ficha"""
        try:
            self.logger.info(f"Procesando campo de input para ficha: {ficha}")
            
            # Verificar que el elemento sea visible e interactuable
            if not input_element.is_displayed():
                self.logger.error("❌ El campo de input no es visible")
                return False
            
            if not input_element.is_enabled():
                self.logger.error("❌ El campo de input no está habilitado")
                return False
            
            # Hacer scroll al elemento
            self.driver.execute_script("arguments[0].scrollIntoView(true);", input_element)
            time.sleep(0.5)
            
            # Hacer clic para enfocar el campo
            self.logger.info("Haciendo clic en el campo de input...")
            input_element.click()
            time.sleep(0.5)
            
            # Limpiar el campo
            self.logger.info("Limpiando el campo...")
            input_element.clear()
            time.sleep(0.5)
            
            # Insertar la ficha
            self.logger.info(f"Insertando ficha: {ficha}")
            input_element.send_keys(str(ficha))
            time.sleep(1)
            
            # Verificar que se insertó correctamente
            inserted_value = input_element.get_attribute('value')
            if inserted_value == str(ficha):
                self.logger.info(f"✓ Ficha '{ficha}' insertada correctamente")
                time.sleep(1)
                return True
            else:
                self.logger.error(f"❌ Error: Se esperaba '{ficha}' pero se insertó '{inserted_value}'")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ Error al procesar el campo de input: {e}")
            return False

    def process_single_ficha(self, ficha, timeout=10):
        """Flujo completo: Primera Opción -> Consultar ficha -> Insertar ficha -> Consultar -> Agregar -> Consultar aspirantes."""
        try:
            self.logger.info(f"\n{'='*50}")
            self.logger.info(f"PROCESANDO FICHA: {ficha}")
            self.logger.info(f"{'='*50}")

            # 1. Seleccionar Primera Opción
            ruta_iframes = self.seleccionar_primera_opcion()
            if ruta_iframes is None:
                self.logger.error("❌ No se pudo seleccionar 'Primera Opción'")
                return False

            # 2. Volver al mismo iframe y hacer clic en Consultar ficha
            self.driver.switch_to.default_content()
            for idx, pos in enumerate(ruta_iframes):
                if idx == 0:
                    outer = self.driver.find_elements(By.TAG_NAME, "iframe")[pos]
                    self.driver.switch_to.frame(outer)
                elif idx == 1:
                    inner = self.driver.find_elements(By.TAG_NAME, "iframe")[pos]
                    self.driver.switch_to.frame(inner)

            if not self.try_click_ficha_button_in_current_frame():
                self.logger.error(f"❌ No se pudo hacer clic en 'Consultar ficha' para ficha {ficha}")
                return False

            self.logger.info("Esperando formulario para ficha...")
            if not self.wait_for_form_and_insert_ficha(ficha):
                self.logger.error(f"❌ No se pudo insertar ficha {ficha}")
                return False

            self.logger.info(f"✓ Ficha {ficha} procesada completamente")
            return True

        except Exception as e:
            self.logger.error(f"❌ Error general al procesar ficha {ficha}: {e}")
            return False

    def run_automation(self, excel_path):
        """Ejecuta la automatización completa"""
        try:
            # Leer fichas del Excel
            fichas = self.read_excel_fichas(excel_path)
            if not fichas:
                return
            
            # Tomar solo las primeras 3 fichas para testing
            fichas = fichas[:1]  # Solo 1 ficha para testing completo
            
            # Navegar al sitio (una sola vez)
            if not self.navigate_to_sena():
                self.logger.error("❌ Error en la navegación inicial")
                return
            
            # Procesar cada ficha
            successful = 0
            failed = 0
            
            for i, ficha in enumerate(fichas):
                self.logger.info(f"\n--- Procesando ficha {i+1} de {len(fichas)} ---")
                
                if self.process_single_ficha(ficha):
                    successful += 1
                else:
                    failed += 1
                
                # Pausa entre fichas
                if i < len(fichas) - 1:  # No pausar después de la última
                    self.logger.info("Esperando antes de la siguiente ficha...")
                    time.sleep(5)  # Tiempo para que se procesen los resultados
            
            self.logger.info(f"\n{'='*50}")
            self.logger.info(f"AUTOMATIZACIÓN COMPLETADA")
            self.logger.info(f"Exitosas: {successful}")
            self.logger.info(f"Fallidas: {failed}")
            self.logger.info(f"{'='*50}")
            
        except Exception as e:
            self.logger.error(f"Error en la automatización: {e}")
        finally:
            # Cerrar archivo HTML
            try:
                with open(self.html_dump_file, 'a', encoding='utf-8') as f:
                    f.write("</body></html>")
            except:
                pass
            
            input("Presiona Enter para cerrar el navegador...")
            self.driver.quit()

def main():
    excel_path = "fichas_sena.xlsx"
    automation = SenaAutomation()
    automation.run_automation(excel_path)

if __name__ == "__main__":
    main()