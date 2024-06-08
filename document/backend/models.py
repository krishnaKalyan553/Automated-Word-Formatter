from django.db import models

class ProcessedFile(models.Model):
    title = models.CharField(max_length=100)
    processed_document = models.FileField(upload_to="processed_documents/")

    def __str__(self):
        return self.title