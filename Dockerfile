ARG IMAGE=intersystemsdc/iris-community:latest
FROM $IMAGE

USER root

# openpyxl into IRIS's embedded Python - see README.md "Requirements". Installed as
# root before ${ISC_PACKAGE_MGRUSER} takes over below, since IRIS's python site-packages
# dir isn't writable by that user.
RUN ${ISC_PACKAGE_INSTALLDIR}/bin/irispython -m pip install openpyxl

WORKDIR /opt/irisapp
RUN chown ${ISC_PACKAGE_MGRUSER}:${ISC_PACKAGE_IRISGROUP} /opt/irisapp

USER ${ISC_PACKAGE_MGRUSER}

COPY module.xml .
COPY src src
COPY iris.script /tmp/iris.script

# A landing page at /csp/user/index.html - static files bypass this app's normal CSP
# login requirement (ServeFiles=2 on /csp/user), so it's the one URL a docker-run user
# can open with no login. IRIS's own web server has no directory-index/default-document
# support (confirmed live: a bare "/" or "/csp/user/" request 404s even with an
# index.html/index.csp physically present - only an exact filename resolves), so
# "http://localhost:52773/csp/user/index.html" is the shortest URL actually achievable,
# not truly the bare host:port.
COPY --chown=${ISC_PACKAGE_MGRUSER}:${ISC_PACKAGE_IRISGROUP} docker/landing/ ${ISC_PACKAGE_INSTALLDIR}/csp/user/

# Installs Samples-BI (sample cubes/dashboards to export) and this module (via ZPM's
# directory-load form, from the copied source rather than the public registry) into
# the default USER namespace, then enables DeepSee there.
RUN iris start IRIS \
    && iris session IRIS < /tmp/iris.script \
    && iris stop IRIS quietly

# Everything docker-entrypoint.sh's own first-boot bootstrap (docker_enable_callin/
# docker_setup_namespace/docker_setup_username) would do is already done above at build
# time - and confirmed live, that bootstrap's docker_setup_namespace step (an embedded-
# Python dbapi connection via irissqlcli) reliably crashes the container on first `docker
# run` on at least one real host/image combination ("Cannot call an iris.package wrapper
# ... dbapi.connect", reproduced against a completely stock, unmodified upstream image -
# not something introduced by this Dockerfile). Pre-creating the marker file that same
# bootstrap writes on success (docker_setup_env checks IRIS_INIT via "[ -s
# "$INSTALLDIR/iris.init" ]") makes the entrypoint see first-boot setup as already done
# and skip straight to "Already initialized, skipping" instead.
RUN date > ${ISC_PACKAGE_INSTALLDIR}/iris.init
