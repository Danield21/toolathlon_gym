import argparse
import json
import logging
import sys
import os
import tempfile
from mcp.server.fastmcp import FastMCP
from .config import config_manager
from .services import EmailService, FolderService, SearchService, DraftService
from .backends.pg_backend import PgIMAPBackend, PgSMTPBackend, PgDraftBackend
from .backends import FileBackend
from .models.config import EmailConfig
from .tools import register_email_tools, register_folder_tools, register_management_tools


def setup_logging(debug: bool = False):
    """Setup logging configuration"""
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stderr)
        ]
    )


def _make_default_email_config() -> EmailConfig:
    """Create a default EmailConfig for PG mode (no real IMAP/SMTP needed)."""
    return EmailConfig(
        email=os.environ.get("EMAIL_ADDRESS", "user@example.com"),
        name=os.environ.get("EMAIL_NAME", "PG Email User"),
        imap_server="localhost",
        imap_port=993,
        smtp_server="localhost",
        smtp_port=587,
        password="unused",
    )


def _prepare_pg_config_file(config_file: str, logger) -> str:
    """Create a PG-safe email config without requiring real credentials.

    Toolathlon email tasks sometimes provide only an address and display name.
    That is sufficient for the PostgreSQL-backed IMAP/SMTP implementation, but
    the shared config loader still validates a password. Normalize every PG
    launch through a temporary config with an explicit placeholder password.
    """
    account = {}
    if config_file and os.path.exists(config_file):
        with open(config_file, "r", encoding="utf-8") as fh:
            config_data = json.load(fh)
        if isinstance(config_data, list):
            account = config_data[0] if config_data else {}
        elif isinstance(config_data, dict):
            account = config_data
        else:
            raise ValueError("Email configuration must be an object or list")
        if not isinstance(account, dict):
            raise ValueError("Email account configuration must be an object")
        logger.info(
            "Using task email identity from '%s' in PG-only mode", config_file
        )
    elif config_file:
        logger.info(
            "Config file '%s' not found - running in PG-only mode", config_file
        )
    else:
        logger.info("No config file specified - running in PG-only mode")

    pg_config = {
        "email": account.get("email")
        or os.environ.get("EMAIL_ADDRESS", "user@example.com"),
        "name": account.get("name")
        or os.environ.get("EMAIL_NAME", "PG Email User"),
        "imap_server": "localhost",
        "imap_port": account.get("imap_port", 993),
        "smtp_server": "localhost",
        "smtp_port": account.get("smtp_port", 587),
        "use_ssl": account.get("use_ssl", True),
        "use_starttls": account.get("use_starttls", True),
        "password": account.get("password") or "unused",
    }
    fd, temp_config_file = tempfile.mkstemp(
        suffix=".json", prefix="email_pg_cfg_"
    )
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        json.dump(pg_config, fh)
    return temp_config_file


def create_services(email_config):
    """Create service instances using PostgreSQL backends"""
    # Create PG-backed backends
    imap_backend = PgIMAPBackend(email_config)
    smtp_backend = PgSMTPBackend(email_config)

    email_export_path = config_manager.workspace_config.email_export_path if config_manager.workspace_config else None
    attachment_download_path = config_manager.workspace_config.attachment_download_path if config_manager.workspace_config else None
    file_backend = FileBackend(email_export_path, attachment_download_path)

    # Create services, injecting PG backends into EmailService
    email_service = EmailService(email_config)
    # Replace the IMAP/SMTP backends that EmailService created internally
    email_service.imap_backend = imap_backend
    email_service.smtp_backend = smtp_backend

    folder_service = FolderService(imap_backend)
    search_service = SearchService(imap_backend)
    draft_service = DraftService(file_backend)

    return email_service, folder_service, search_service, draft_service


def main():
    """Main function to run the emails MCP server"""
    parser = argparse.ArgumentParser(description='Emails MCP Server')
    parser.add_argument(
        '--attachment_upload_path',
        type=str,
        default=None,
        help='Directory path for attachment uploads (restricts file selection to this path and subdirectories)'
    )
    parser.add_argument(
        '--attachment_download_path',
        type=str,
        default=None,
        help='Directory path for attachment downloads (files will be saved here with unique names)'
    )
    parser.add_argument(
        '--email_export_path',
        type=str,
        default=None,
        help='Directory path for email exports (exports will be saved here with date-based filenames)'
    )
    parser.add_argument(
        '--config_file',
        type=str,
        default=None,
        help='Email configuration file path (optional for PG mode)'
    )
    parser.add_argument(
        '--debug',
        action='store_true',
        help='Enable debug logging'
    )

    args = parser.parse_args()

    # Setup logging
    setup_logging(args.debug)
    logger = logging.getLogger(__name__)

    try:
        # Initialize MCP server
        mcp = FastMCP("emails-mcp")

        # This server always uses PostgreSQL-backed IMAP/SMTP. Preserve the
        # task's email identity, but never require real network credentials.
        _temp_config_file = _prepare_pg_config_file(args.config_file, logger)
        config_file = _temp_config_file

        # Load configuration
        config_manager.load_workspace_config(
            attachment_upload_path=args.attachment_upload_path,
            attachment_download_path=args.attachment_download_path,
            email_export_path=args.email_export_path,
            config_file=config_file
        )

        email_config = config_manager.load_email_config(config_file)
        if not email_config:
            # Fallback: build a default config for PG mode
            email_config = _make_default_email_config()

        logger.info(f"Loaded configuration for: {email_config.email}")

        # Create services (PG-backed)
        email_service, folder_service, search_service, draft_service = create_services(email_config)

        # Register MCP tools
        register_email_tools(mcp, email_service)
        register_folder_tools(mcp, folder_service)
        register_management_tools(mcp, draft_service, email_service)

        logger.info("All MCP tools registered successfully")

        # Log path restrictions if set
        if config_manager.workspace_config:
            config = config_manager.workspace_config
            if config.attachment_upload_path:
                logger.info(f"Attachment uploads restricted to: {config.attachment_upload_path}")
            if config.attachment_download_path:
                logger.info(f"Attachment downloads will be saved to: {config.attachment_download_path}")
            if config.email_export_path:
                logger.info(f"Email exports will be saved to: {config.email_export_path}")

        # Test PG connection on startup
        try:
            imap_ok, smtp_ok = email_service.check_connection()
            if imap_ok and smtp_ok:
                logger.info("All PG email connections verified successfully")
            else:
                logger.warning("PG email connection check returned partial failure")
        except Exception as e:
            logger.warning(f"PG connection test failed: {str(e)}")

        # Start the MCP server
        logger.info("Starting emails MCP server (PG backend)...")
        mcp.run(transport='stdio')

    except KeyboardInterrupt:
        logger.info("Server shutdown requested")
    except Exception as e:
        logger.error(f"Server startup failed: {str(e)}")
        sys.exit(1)
    finally:
        # Cleanup
        try:
            if 'email_service' in locals():
                email_service.cleanup()
        except:
            pass
        # Remove temporary config file if we created one
        try:
            if '_temp_config_file' in locals() and _temp_config_file and os.path.exists(_temp_config_file):
                os.unlink(_temp_config_file)
        except:
            pass


if __name__ == "__main__":
    main()
