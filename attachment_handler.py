import logging

logger = logging.getLogger(__name__)

class AttachmentHandler:
    def process_attachments(self, attachments):
        try:
            for attachment in attachments:
                logger.info(f"get Attachment: {attachment.name}")
                self.save_attachment(attachment)
        except Exception as e:
            logger.error(f"Error processing attachments: {e}")

    def save_attachment(self, attachment):
        try:
            with open(attachment.name, 'wb') as f:
                if attachment.name.endswith('.eml'):
                    f.write(attachment.item.mime_content)
                else:
                    f.write(attachment.content)

        except Exception as e:
            logger.error(f"Error saving attachment: {e}")