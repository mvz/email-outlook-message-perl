FROM perl:5 AS builder

WORKDIR /build
COPY . /build/
RUN cpanm --installdeps --notest --no-man-pages .


FROM perl:5-slim

ENV WORKDIR '/opt/email-outlook-message-perl'
WORKDIR $WORKDIR
COPY . .
COPY --from=builder /usr/local/lib/perl5 /usr/local/lib/perl5

RUN perl Build.PL \
 && ./Build \
 && ./Build test \
 && ./Build install

ENTRYPOINT ["msgconvert"]
