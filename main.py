for i in range(0, len(data), 2)) / (32768.0 * len(data) / 2)))
                    self.audio_level.set(level)
                    
                    try:
                        audio = self.recognizer.listen(source, timeout=1)
                        text = self.recognizer.recognize_google(
                            audio,
                            language=self.lang_var.get().split()[1].strip('()')
                        )
                        
                        self.window.after(0, self.insert_text, text)
                        
                    except sr.WaitTimeoutError:
                        continue
                    except sr.UnknownValueError:
                        continue
                    except sr.RequestError:
                        self.speech_label.configure(
                            text="Could not connect to speech recognition service",
                            text_color="red"
                        )
                        break
                        
        finally:
            self.audio_manager.stop_stream()
            self.audio_level.set(0)
    
    def stop_recording(self):
        """Stop voice recording"""
        self.is_recording = False
        self.speech_label.configure(text="Recording stopped", text_color="gray")
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.save_audio_button.configure(state="normal")
    
    def save_recording(self):
        """Save recorded audio to file"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".wav",
            filetypes=[("Wave files", "*.wav")]
        )
        
        if filename:
            self.audio_manager.save_recording(filename)
            messagebox.showinfo("Success", "Recording saved successfully!")
    
    def insert_text(self, text):
        """Insert recognized text at cursor position"""
        current_pos = self.text_area.index("insert")
        self.text_area.insert(current_pos, f" {text}")
        self.update_stats()
    
    def import_pdf(self):
        """Import text from PDF file"""
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            try:
                text = self.doc_manager.import_pdf(file_path)
                self.text_area.delete("1.0", "end")
                self.text_area.insert("1.0", text)
                self.current_file = file_path
                self.update_stats()
                self.update_toc()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import PDF: {str(e)}")
    
    def import_word(self):
        """Import text from Word document"""
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if file_path:
            try:
                text = self.doc_manager.import_word(file_path)
                self.text_area.delete("1.0", "end")
                self.text_area.insert("1.0", text)
                self.current_file = file_path
                self.update_stats()
                self.update_toc()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import Word document: {str(e)}")
    
    def import_image(self):
        """Import text from image using OCR"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        )
        if file_path:
            try:
                text = self.doc_manager.import_image(file_path)
                self.text_area.delete("1.0", "end")
                self.text_area.insert("1.0", text)
                self.current_file = file_path
                self.update_stats()
                self.update_toc()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import image: {str(e)}")
    
    def import_text(self):
        """Import plain text file"""
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    text = file.read()
                self.text_area.delete("1.0", "end")
                self.text_area.insert("1.0", text)
                self.current_file = file_path
                self.update_stats()
                self.update_toc()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import text file: {str(e)}")
    
    def export_pdf(self):
        """Export text to PDF file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            try:
                text = self.text_area.get("1.0", "end-1c")
                self.doc_manager.export_pdf(text, file_path)
                messagebox.showinfo("Success", "Text exported to PDF successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export PDF: {str(e)}")
    
    def export_word(self):
        """Export text to Word document"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")]
        )
        if file_path:
            try:
                text = self.text_area.get("1.0", "end-1c")
                self.doc_manager.export_word(text, file_path)
                messagebox.showinfo("Success", "Text exported to Word document successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export Word document: {str(e)}")
    
    def export_text(self):
        """Export to plain text file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")]
        )
        if file_path:
            try:
                text = self.text_area.get("1.0", "end-1c")
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(text)
                messagebox.showinfo("Success", "Text exported successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export text: {str(e)}")
    
    def new_document(self):
        """Create new document"""
        if messagebox.askyesno("New Document", "Clear current text?"):
            self.text_area.delete("1.0", "end")
            self.current_file = None
            self.update_stats()
            self.update_toc()
    
    def clear_text(self):
        """Clear all text"""
        if messagebox.askyesno("Clear Text", "Are you sure you want to clear all text?"):
            self.text_area.delete("1.0", "end")
            self.update_stats()
            self.update_toc()
    
    def on_closing(self):
        """Handle window closing"""
        if self.is_recording:
            self.stop_recording()
        self.window.destroy()
    
    def run(self):
        """Start the application"""
        self.window.mainloop()

if __name__ == "__main__":
    app = EnhancedVoiceTypingApp()
    app.run()