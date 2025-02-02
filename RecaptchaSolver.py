import os
import pydub
import random
import speech_recognition
import time
import urllib
from time import sleep

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class RecaptchaSolver:
    def __init__(self, driver: webdriver.Chrome):
        self.driver = driver


    def solve(self):
        sleep(2)
        # Switch to the reCAPTCHA iframe
        self._switch_to_recaptcha_iframe()
        # Click on the reCAPTCHA checkbox
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.rc-anchor-content'))
        ).click()
        print("Clicked on the reCAPTCHA checkbox")

        # Check if the CAPTCHA is solved
        if self.isSolved():
            return

        sleep(2)
        # Switch to the challenge iframe to handle audio CAPTCHA
        self.driver.switch_to.default_content()  # Go back to the main document
        self._switch_to_challenge_iframe()
        sleep(1)

        # Click on the audio button
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'recaptcha-audio-button'))
        ).click()
        print("Clicked on the audio button")
        time.sleep(1)

        # Get the audio source and recognize speech from it
        key = self._solve_audio_captcha(play_audio=True)

        # Input the key
        self.driver.find_element(By.ID, 'audio-response').send_keys(key.lower())
        time.sleep(1)

        # Click on verify button
        self.driver.find_element(By.ID, 'recaptcha-verify-button').click()
        time.sleep(5)

        # Check if the captcha is solved
        self.driver.switch_to.default_content()  # Go back to the main document
        self._switch_to_recaptcha_iframe()
        if self.isSolved():
            return
        else:
            raise Exception("Failed to solve the captcha")

    def isSolved(self):
        try:
            checkbox = self.driver.find_element(By.CSS_SELECTOR, '.recaptcha-checkbox-checkmark')
            style = checkbox.get_dom_attribute('style')
            return style is not None
        except Exception as e:
            print(e)
            return False

    def _switch_to_recaptcha_iframe(self):
        # Wait for the outer iframe to load
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[title='reCAPTCHA']"))
        )
        outer_iframe = self.driver.find_element(By.CSS_SELECTOR, "iframe[title='reCAPTCHA']")
        self.driver.switch_to.frame(outer_iframe)
        print("Switched to outer iframe")

    def _switch_to_challenge_iframe(self):
        # Wait for the challenge iframe to load
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@title, 'recaptcha')]"))
        )
        challenge_iframe = self.driver.find_element(By.XPATH, "//iframe[contains(@title, 'recaptcha')]")
        self.driver.switch_to.frame(challenge_iframe)
        print("Switched to challenge iframe")

    def _solve_audio_captcha(self, play_audio=False):
        max_retry = 3
        while max_retry:
            try:
                if play_audio:
                    self._play_audio()
                # Get the initial audio source
                initial_src = self.driver.find_element(By.ID, 'audio-source').get_attribute('src')
                # Attempt to get the key from the audio source
                return self._get_key_from_audio_source(src=initial_src)
            except speech_recognition.exceptions.UnknownValueError:
                print("Speech recognition failed, retrying...")

                # Click on the reload button to get a new challenge
                self.driver.find_element(
                    By.XPATH,
                    "//button[contains(@class, 'rc-button-reload') and contains(@title, 'new challenge')]"
                ).click()

                # Wait until the audio source changes
                WebDriverWait(self.driver, 10).until(
                    lambda driver: driver.find_element(By.ID, 'audio-source').get_attribute('src') != initial_src
                )

                print("Audio source reloaded.")
                max_retry -= 1
                time.sleep(1)  # Optional: Brief pause after ensuring the source is reloaded

        raise Exception("Failed to solve audio CAPTCHA after maximum retries")

    def _get_key_from_audio_source(self, src):
        # Download the audio to the temp folder
        path_to_mp3 = os.path.normpath(
            os.path.join(self._temp_dir() + str(random.randrange(1, 1000)) + ".mp3")
        )
        path_to_wav = os.path.normpath(
            os.path.join(self._temp_dir() + str(random.randrange(1, 1000)) + ".wav")
        )
        urllib.request.urlretrieve(src, path_to_mp3)

        # Convert mp3 to wav
        sound = pydub.AudioSegment.from_mp3(path_to_mp3)
        sound.export(path_to_wav, format="wav")

        # Recognize speech
        print("Recognizing Speech start")
        sample_audio = speech_recognition.AudioFile(path_to_wav)
        r = speech_recognition.Recognizer()
        with sample_audio as source:
            audio = r.record(source)

        # Recognize the audio
        return r.recognize_google(audio)

    def _play_audio(self):
        self.driver.find_element(
            By.XPATH,
            "//button[contains(@class, 'rc-button-default') and text()='PLAY' and @aria-labelledby='audio-instructions rc-response-label']",
        ).click()
        time.sleep(5)


    def _temp_dir(self):
        return os.getenv("TEMP") if os.name == "nt" else "/tmp/"

    def _random_scroll(self):
        element = self.driver.find_element(By.TAG_NAME, "body")
        element.send_keys(Keys.PAGE_DOWN)
        time.sleep(1)
        element.send_keys(Keys.PAGE_UP)
        time.sleep(1)