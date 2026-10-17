#!/bin/env python3
# pylint: disable=no-value-for-parameter

"""

XSD Downloader

"""

import os
import re
import hashlib
import ctypes
import time
import logging
import urllib.request
from urllib.parse import urljoin, urlparse, unquote

from Settings import Settings

settings = Settings()

XSD_DIR = f"{settings.get_docs_path()}/cache"
DOWNLOAD_TIMEOUT = 30


def schema_cache_dir(schema_url: str) -> str:
    """
    Every root schema gets its own cache folder, so labels from different
    schema versions (which reuse the same file names) never get mixed
    """
    url_hash = hashlib.sha256(schema_url.encode()).hexdigest()[:16]
    return f"{XSD_DIR}/{url_hash}"


def url_to_path(url: str) -> str:
    """
    Turns a URL to an XSD into a filesystem path
    eg "http://www.example.come/a/b/c.xsd" -> "./www.example.come/a/b/c.xsd"
    """

    parsed = urlparse(url)
    path = parsed.netloc + parsed.path + parsed.query
    path = parsed.path
    if parsed.query and "/" in parsed.query:
        path = path + parsed.query[parsed.query.rfind("/") :]
    return path


def plain_filename(url: str) -> str:
    last_slash = url.rfind(r'/')
    last_xsd = url.rfind('.xsd')

    if last_xsd == -1:
        return f'{time.time()}.xsd'

    return url[last_slash+1:last_xsd+4]


def localize_links(text: str, filename_complete: str) -> str:
    """
    Similar to wget's --convert-links, this converts the schemaLocation
    links to be useable on local filesystem
    """

    # only converting the url-based schemaLocations here
    # eg, some will schemaLocation="../abc.xsd"
    schema_locations = re.findall('schemaLocation="(http[^"]*)"', text)

    for schema_location in schema_locations:
        path_url = url_to_path(schema_location)
        rel_path = os.path.relpath(
            os.path.dirname(path_url), os.path.dirname(filename_complete)
        )
        base_name = os.path.basename(path_url)
        text = text.replace(schema_location, rel_path + "/" + base_name)
    return text


def save_file(url: str, text: str, file_name: str) -> None:
    """
    Save the XSD `text` data to file path decided by `url`
    also creates the directory structure if it doesn't exist
    """

    dir_name = os.path.dirname(file_name)

    # create directory structure
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
        if os.name == "nt":
            FILE_ATTRIBUTE_HIDDEN = 0x02
            ctypes.windll.kernel32.SetFileAttributesW(dir_name, FILE_ATTRIBUTE_HIDDEN)

    #     return last_path
    with open(file_name, "w", encoding="utf_8_sig") as f:
        # text_localized = localize_links(text, filename_complete)
        f.write(text)
        f.close()


def download_xml_url(url: str) -> str:
    """
    Fetches URL, returns text of the document
    """

    urlopen = urllib.request.urlopen(url, timeout=DOWNLOAD_TIMEOUT)
    response = urlopen.read()
    headers = dict(urlopen.headers.items())
    try:
        content_disposition = headers['content-disposition']
        content_disposition = content_disposition.replace('attachment; filename*=UTF-8\'\'', '')
        filename = unquote(urlparse(content_disposition).path)
    except KeyError:
        filename = None

    return response.decode("utf-8"), filename


def download_schema(url: str) -> None:
    """
    Calls the recursive function recursive_get_schema_locations()
    """

    # list of the URLs that have been downloaded already
    downloaded_urls = []
    cache_dir = schema_cache_dir(url)

    def recursive_get_schema_locations(url: str, referring_url: str) -> None:
        """
        Recursive function that is the heart of the script,
        stops when... TODO
        """
        file_path_from_url = cache_dir + "/" + plain_filename(url)
        # dont download same link twice
        if url not in downloaded_urls and not os.path.exists(file_path_from_url):
            downloaded_urls.append(url)
            try:
                xsd_data, header_filename = download_xml_url(url)

            except (urllib.error.URLError, TimeoutError) as e:
                logging.error(
                    "ERROR loading {} , referenced in {} REASON: {} ".format(
                        url, referring_url, getattr(e, "reason", e)
                    )
                )
                return
            if header_filename is not None and header_filename != plain_filename(url):
                file_path_from_url = cache_dir + "/" + header_filename
            # all the XSDs linked from this file via schemaLocation
            schema_locations = re.findall('schemaLocation="([^"]*)"', xsd_data)

            # write this file in the directory structure
            save_file(url, xsd_data, file_path_from_url)

            # iterate through all schemaLocation URLs
            for schema_location in schema_locations:
                # urljoin translates relative URLs in schemaLocation to
                # absolute

                # eg:
                # >>> urljoin('http://example.com/1/2/3','../../index.html')
                # turns into: http://example.com/index.html

                line = urljoin(url, schema_location)
                logging.debug(line)
                recursive_get_schema_locations(line, url)

    recursive_get_schema_locations(url, url)

if __name__ == "__main__":
    pass